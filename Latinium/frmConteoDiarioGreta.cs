using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmConteoDiarioGreta.
	/// </summary>
	public class frmConteoDiarioGreta : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label55;
		private System.Windows.Forms.Label label53;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label lblContador;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSubGrupo;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtArticulo;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Label lblVendedor;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.ContextMenu mnuImprimir;
		private System.Windows.Forms.MenuItem mnuImprimirInventario;
		private System.Windows.Forms.MenuItem mnuImprimirInforme;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridConteo;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txidConteo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private System.Windows.Forms.Button btnGenerar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmConteoDiarioGreta()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCategoria");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Categoria");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCategoria");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Categoria");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmConteoDiarioGreta));
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSubgrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleConteoDiarioInventario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoM");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Existencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Exibicion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Conteo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Diferencias");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleConteoDiarioInventario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoM");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Existencia");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Exibicion");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Conteo");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Diferencias");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Existencia", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Existencia", 6, true);
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Conteo", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Conteo", 9, true);
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Exibicion", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Exibicion", 7, true);
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Bodega", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Bodega", 8, true);
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Diferencias", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Diferencias", 11, true);
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			this.label7 = new System.Windows.Forms.Label();
			this.label55 = new System.Windows.Forms.Label();
			this.label53 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.lblContador = new System.Windows.Forms.Label();
			this.cmbSubGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblEstado = new System.Windows.Forms.Label();
			this.txtArticulo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnVer = new System.Windows.Forms.Button();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblVendedor = new System.Windows.Forms.Label();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.mnuImprimir = new System.Windows.Forms.ContextMenu();
			this.mnuImprimirInventario = new System.Windows.Forms.MenuItem();
			this.mnuImprimirInforme = new System.Windows.Forms.MenuItem();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.uGridConteo = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.txidConteo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnGenerar = new System.Windows.Forms.Button();
			this.label2 = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridConteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txidConteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			this.SuspendLayout();
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Enabled = false;
			this.label7.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.Location = new System.Drawing.Point(8, 112);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(37, 17);
			this.label7.TabIndex = 398;
			this.label7.Text = "Buscar";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label55
			// 
			this.label55.AutoSize = true;
			this.label55.BackColor = System.Drawing.Color.Transparent;
			this.label55.Enabled = false;
			this.label55.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label55.Location = new System.Drawing.Point(8, 48);
			this.label55.Name = "label55";
			this.label55.Size = new System.Drawing.Size(41, 17);
			this.label55.TabIndex = 397;
			this.label55.Text = "Bodega";
			this.label55.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label53
			// 
			this.label53.AutoSize = true;
			this.label53.BackColor = System.Drawing.Color.Transparent;
			this.label53.Enabled = false;
			this.label53.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label53.Location = new System.Drawing.Point(8, 16);
			this.label53.Name = "label53";
			this.label53.Size = new System.Drawing.Size(34, 17);
			this.label53.TabIndex = 396;
			this.label53.Text = "Fecha";
			this.label53.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Enabled = false;
			this.label12.Location = new System.Drawing.Point(8, 80);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(53, 16);
			this.label12.TabIndex = 401;
			this.label12.Text = "Categoria";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Enabled = false;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.lblContador.Location = new System.Drawing.Point(448, 112);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 17);
			this.lblContador.TabIndex = 402;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// cmbSubGrupo
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSubGrupo.Appearance = appearance1;
			this.cmbSubGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSubGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSubGrupo.DataSource = this.ultraDataSource3;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 264;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 264;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbSubGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbSubGrupo.DisplayMember = "Categoria";
			this.cmbSubGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSubGrupo.Enabled = false;
			this.cmbSubGrupo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSubGrupo.Location = new System.Drawing.Point(80, 77);
			this.cmbSubGrupo.Name = "cmbSubGrupo";
			this.cmbSubGrupo.Size = new System.Drawing.Size(264, 22);
			this.cmbSubGrupo.TabIndex = 400;
			this.cmbSubGrupo.ValueMember = "idCategoria";
			this.cmbSubGrupo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbSubGrupo_KeyDown);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.lblEstado.Location = new System.Drawing.Point(496, 456);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 23);
			this.lblEstado.TabIndex = 399;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtArticulo
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtArticulo.Appearance = appearance2;
			this.txtArticulo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtArticulo.Enabled = false;
			this.txtArticulo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtArticulo.Location = new System.Drawing.Point(80, 109);
			this.txtArticulo.Name = "txtArticulo";
			this.txtArticulo.Size = new System.Drawing.Size(264, 22);
			this.txtArticulo.TabIndex = 394;
			this.txtArticulo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtArticulo_KeyDown);
			// 
			// btnVer
			// 
			this.btnVer.Enabled = false;
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(360, 108);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 395;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 180;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(80, 45);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(264, 22);
			this.cmbBodega.TabIndex = 392;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyDown);
			this.cmbBodega.Validating += new System.ComponentModel.CancelEventHandler(this.cmbBodega_Validating);
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			// 
			// dtFecha
			// 
			appearance3.FontData.Name = "Tahoma";
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance3;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(80, 14);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 393;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.Validating += new System.ComponentModel.CancelEventHandler(this.dtFecha_Validating);
			this.dtFecha.Validated += new System.EventHandler(this.dtFecha_Validated);
			// 
			// lblVendedor
			// 
			this.lblVendedor.AutoSize = true;
			this.lblVendedor.Enabled = false;
			this.lblVendedor.Location = new System.Drawing.Point(360, 48);
			this.lblVendedor.Name = "lblVendedor";
			this.lblVendedor.Size = new System.Drawing.Size(42, 16);
			this.lblVendedor.TabIndex = 404;
			this.lblVendedor.Text = "Realiza";
			// 
			// cmbVendedor
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVendedor.Appearance = appearance4;
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn6.Header.VisiblePosition = 0;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn7.Width = 335;
			ultraGridColumn8.Header.VisiblePosition = 2;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbVendedor.Enabled = false;
			this.cmbVendedor.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(424, 48);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(264, 22);
			this.cmbVendedor.TabIndex = 403;
			this.cmbVendedor.ValueMember = "idPersonal";
			this.cmbVendedor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbVendedor_KeyDown);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// ultraDataSource1
			// 
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn14.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
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
			// mnuImprimir
			// 
			this.mnuImprimir.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																								this.mnuImprimirInventario,
																																								this.mnuImprimirInforme});
			// 
			// mnuImprimirInventario
			// 
			this.mnuImprimirInventario.Enabled = false;
			this.mnuImprimirInventario.Index = 0;
			this.mnuImprimirInventario.Text = "Inventario";
			// 
			// mnuImprimirInforme
			// 
			this.mnuImprimirInforme.Enabled = false;
			this.mnuImprimirInforme.Index = 1;
			this.mnuImprimirInforme.Text = "Informe";
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
			// uGridConteo
			// 
			this.uGridConteo.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridConteo.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridConteo.DataSource = this.ultraDataSource1;
			appearance5.BackColor = System.Drawing.Color.White;
			this.uGridConteo.DisplayLayout.Appearance = appearance5;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 2;
			ultraGridColumn11.Width = 170;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.Caption = "Artículo";
			ultraGridColumn12.Header.VisiblePosition = 3;
			ultraGridColumn12.Width = 230;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn13.Header.VisiblePosition = 8;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 180;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn14.Header.Caption = "Estado";
			ultraGridColumn14.Header.VisiblePosition = 4;
			ultraGridColumn14.Width = 100;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance6;
			ultraGridColumn15.Format = "#,##0";
			ultraGridColumn15.Header.VisiblePosition = 5;
			ultraGridColumn15.Width = 75;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance7;
			ultraGridColumn16.Format = "#,##0";
			ultraGridColumn16.Header.Caption = "Exhibicion";
			ultraGridColumn16.Header.VisiblePosition = 9;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn16.Width = 75;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance8;
			ultraGridColumn17.Format = "#,##0";
			ultraGridColumn17.Header.VisiblePosition = 10;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn17.Width = 75;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance9;
			ultraGridColumn18.Format = "#,##0";
			ultraGridColumn18.Header.VisiblePosition = 6;
			ultraGridColumn18.Width = 50;
			ultraGridColumn19.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn19.Header.VisiblePosition = 11;
			ultraGridColumn19.Width = 200;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance10;
			ultraGridColumn20.Header.VisiblePosition = 7;
			ultraGridBand4.Columns.AddRange(new object[] {
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
																										 ultraGridColumn20});
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance11;
			summarySettings1.DisplayFormat = "{0: #,##0}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance12;
			summarySettings2.DisplayFormat = "{0: #,##0}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance13;
			summarySettings3.DisplayFormat = "{0: #,##0}";
			summarySettings3.Hidden = true;
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance14;
			summarySettings4.DisplayFormat = "{0: #,##0}";
			summarySettings4.Hidden = true;
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance15;
			summarySettings5.DisplayFormat = "{0: #,##0}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand4.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4,
																																															summarySettings5});
			ultraGridBand4.SummaryFooterCaption = "";
			this.uGridConteo.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			appearance16.ForeColor = System.Drawing.Color.Black;
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridConteo.DisplayLayout.Override.ActiveRowAppearance = appearance16;
			this.uGridConteo.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridConteo.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridConteo.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance17.BackColor = System.Drawing.Color.Transparent;
			this.uGridConteo.DisplayLayout.Override.CardAreaAppearance = appearance17;
			appearance18.ForeColor = System.Drawing.Color.Black;
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridConteo.DisplayLayout.Override.CellAppearance = appearance18;
			appearance19.ForeColor = System.Drawing.Color.Black;
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridConteo.DisplayLayout.Override.GroupByRowAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance20.FontData.BoldAsString = "True";
			appearance20.FontData.Name = "Arial";
			appearance20.FontData.SizeInPoints = 8.5F;
			appearance20.ForeColor = System.Drawing.Color.White;
			appearance20.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridConteo.DisplayLayout.Override.HeaderAppearance = appearance20;
			this.uGridConteo.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance21.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance21.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridConteo.DisplayLayout.Override.RowAlternateAppearance = appearance21;
			appearance22.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance22.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridConteo.DisplayLayout.Override.RowSelectorAppearance = appearance22;
			appearance23.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance23.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridConteo.DisplayLayout.Override.SelectedRowAppearance = appearance23;
			this.uGridConteo.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridConteo.Location = new System.Drawing.Point(8, 144);
			this.uGridConteo.Name = "uGridConteo";
			this.uGridConteo.Size = new System.Drawing.Size(936, 224);
			this.uGridConteo.TabIndex = 405;
			this.uGridConteo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridConteo_KeyDown);
			this.uGridConteo.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridConteo_AfterCellUpdate);
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(248, 496);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 413;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 496);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 412;
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
			this.btnConsultar.Location = new System.Drawing.Point(88, 496);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 411;
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
			this.btnEditar.Location = new System.Drawing.Point(168, 496);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 410;
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
			this.btnGuardar.Location = new System.Drawing.Point(336, 496);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 409;
			this.btnGuardar.Text = "&Cerrar";
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
			this.btnAnular.Location = new System.Drawing.Point(672, 496);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 408;
			this.btnAnular.Text = "&Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Visible = false;
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(416, 496);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 407;
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
			this.btnSalir.Location = new System.Drawing.Point(496, 496);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 406;
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
			this.groupBox1.Location = new System.Drawing.Point(8, 480);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(936, 8);
			this.groupBox1.TabIndex = 414;
			this.groupBox1.TabStop = false;
			// 
			// txidConteo
			// 
			this.txidConteo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txidConteo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txidConteo.Enabled = false;
			this.txidConteo.Location = new System.Drawing.Point(872, 496);
			this.txidConteo.Name = "txidConteo";
			this.txidConteo.PromptChar = ' ';
			this.txidConteo.Size = new System.Drawing.Size(72, 21);
			this.txidConteo.TabIndex = 416;
			this.txidConteo.Visible = false;
			// 
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(760, 496);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(88, 21);
			this.txtEstado.TabIndex = 415;
			this.txtEstado.Visible = false;
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(8, 384);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(33, 17);
			this.label1.TabIndex = 418;
			this.label1.Text = "Notas";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNotas
			// 
			this.txtNotas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance24;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Enabled = false;
			this.txtNotas.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNotas.Location = new System.Drawing.Point(8, 408);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(472, 66);
			this.txtNotas.TabIndex = 417;
			// 
			// btnGenerar
			// 
			this.btnGenerar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGenerar.CausesValidation = false;
			this.btnGenerar.Enabled = false;
			this.btnGenerar.Image = ((System.Drawing.Image)(resources.GetObject("btnGenerar.Image")));
			this.btnGenerar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGenerar.Location = new System.Drawing.Point(360, 80);
			this.btnGenerar.Name = "btnGenerar";
			this.btnGenerar.Size = new System.Drawing.Size(77, 23);
			this.btnGenerar.TabIndex = 419;
			this.btnGenerar.Text = "Generar";
			this.btnGenerar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGenerar.Click += new System.EventHandler(this.btnGenerar_Click);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(768, 16);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(44, 17);
			this.label2.TabIndex = 421;
			this.label2.Text = "Número";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumero
			// 
			appearance25.ForeColor = System.Drawing.Color.Black;
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance25;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(832, 14);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(112, 21);
			this.txtNumero.TabIndex = 420;
			// 
			// frmConteoDiarioGreta
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(952, 526);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblVendedor);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label55);
			this.Controls.Add(this.label53);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.btnGenerar);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.txidConteo);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.uGridConteo);
			this.Controls.Add(this.cmbVendedor);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.cmbSubGrupo);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtArticulo);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.dtFecha);
			this.KeyPreview = true;
			this.Name = "frmConteoDiarioGreta";
			this.Text = "frmConteoDiarioGreta";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmConteoDiarioGreta_KeyDown);
			this.Load += new System.EventHandler(this.frmConteoDiarioGreta_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridConteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txidConteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		
		}
		bool bNuevo = false;
		bool bEdicion = false;
		private Acceso miAcceso;
		bool bEliminaAlValidar = false;		
		int idBloqueaTransacciones = 0;
		int iBodegaPredef = 0;
		private bool bActualiza = true;
		bool bCombo = false;

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void frmConteoDiarioGreta_Load(object sender, System.EventArgs e)
		{

			string sSQL = string.Format("Exec ListaBodegas {0}, 43", MenuLatinium.IdUsuario);
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);		
			iBodegaPredef = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			string sSQLs = string.Format("Select idCategoria, Categoria From ArticuloCategoria Where Activo = 1 Order By Categoria");
			this.cmbSubGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQLs);

		
			this.cmbVendedor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format(" Select idPersonal, (n.ApellidoPaterno + ' ' + n.ApellidoMaterno + ' ' + n.PrimerNombre + ' ' + n.SegundoNombre) As Nombre, n.Codigo From Nomina n Where  ISNULL((Select Top 1 FechaSalida From NominaEntrada Where idNomina = n.idNomina Order By FechaEntrada Desc), '20301231') > '{0}' Order By Nombre	", 
				Convert.ToDateTime(DateTime.Today).ToString("yyyyMMdd")));
			
			
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			if (bEdicion && !bNuevo)
			{
				string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 3", idBloqueaTransacciones, (int)this.txidConteo.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLBloqueo, true);
			}

			bEdicion = false;
			bNuevo = false;

			this.btnNuevo.Focus();
			this.txtNumero.Text ="";

			this.dtFecha.Value = System.DBNull.Value;
			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;
			this.cmbBodega.Value = System.DBNull.Value;
			this.cmbSubGrupo.Value = System.DBNull.Value;
			this.cmbVendedor.Value = System.DBNull.Value;
			this.txtArticulo.Text = "";

			this.lblContador.Text  = "";
			this.lblEstado.Text ="";
		  
			this.txidConteo.Value = 0;
			this.txtEstado.Value = 0;

			this.txtNotas.Text ="";
			this.uGridConteo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec [ListaConteoDiarioInventarioGreta] 0,'',0");

			this.dtFecha.Enabled = false;
			this.cmbBodega.Enabled = false;
			this.cmbSubGrupo.Enabled = false;
			this.cmbVendedor.Enabled = false;
			this.txtArticulo.Enabled = false;
			this.txtNotas.Enabled = false;		
						FuncionesProcedimientos.EstadoGrids(false, this.uGridConteo);
//
//			if (miAcceso.Nuevo)
				this.btnNuevo.Enabled = true;
//			if (miAcceso.Buscar)
				this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.btnGenerar.Enabled = false;

			idBloqueaTransacciones = 0;
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);						
			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);
				
			bNuevo = true;
			bEdicion = true;

			if (iBodegaPredef > 0) this.cmbBodega.Value = iBodegaPredef;
			else this.cmbBodega.Enabled = true;	

			this.uGridConteo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListaConteoDiarioInventario 0,'',0");

			this.cmbBodega.Enabled = true;
			this.cmbSubGrupo.Enabled = true;
			this.cmbVendedor.Enabled = true;
			this.txtArticulo.Enabled = true;
			this.txtNotas.Enabled = true;
			this.btnVer.Enabled = true;
	
			this.btnGenerar.Enabled = true;

				FuncionesProcedimientos.EstadoGrids(true, this.uGridConteo);

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;
	
			
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

		private void btnGenerar_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Ingrese la Fecha del Inventario", true, "BODEGA", cdsSeteoF)) return;
			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione la Bodega ")) return;
			if (!Validacion.vbComboVacio(this.cmbVendedor, "Seleccione el Responsable del Inventario")) return;

      int idVerificar = 0;

//			   idVerificar = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select  COUNT (*) from ConteoDiarioInventario where Bodega = {0} and CONVERT (dATE, Fecha) = CONVERT (dATE, '{1}') AND Estado IN (1,2)"
//					 ,(int)this.cmbBodega.Value), Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));	

			idVerificar = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select  COUNT (*) from ConteoDiarioInventario where Bodega = {0} and CONVERT (dATE, Fecha) = CONVERT (dATE, '{1}') AND Estado IN (1,2)", (int)this.cmbBodega.Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));


//	
			if (idVerificar >0)
			{
				MessageBox.Show("Ya se genero un conteo de Inventario", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);	
				return ;

			}

			if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de generar el conteo?\n\nEste proceso no puede ser revertido", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				string ssSQL = string.Format("Exec NumeracionLocales 72, {0}, 0", (int)this.cmbBodega.Value);
				this.txtNumero.Text = Funcion.sEscalarSQL(cdsSeteoF, ssSQL);	

				string sSQL = string.Format("Exec [GenerarConteoDiarioInventarioGreta] {0}, {1}, '{2}', 1, '', {3}, '{4}'",(int) this.txidConteo.Value, (int)this.cmbBodega.Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int) this.cmbVendedor.Value, (string)this.txtNumero.Text);
				this.txidConteo.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);	

				MessageBox.Show("Conteo de inventario generado correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				
        
				this.cmbBodega.Enabled = false;
				this.cmbVendedor.Enabled = false;
				this.btnGenerar.Enabled = false;
				this.Consulta();
			}
		}

		private void ListaConteo()
		{
			int idSubgrupo = 0;
			if (this.cmbSubGrupo.ActiveRow != null) idSubgrupo = (int)this.cmbSubGrupo.Value;

			string sSQL = string.Format("Exec ListaConteoDiarioInventarioGreta {0}, '{1}', {2}", (int) this.txidConteo.Value, this.txtArticulo.Text.ToString(), idSubgrupo);
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridConteo);

			this.lblContador.Text = this.uGridConteo.Rows.Count + " REGISTROS ENCONTRADOS";
		}

		private void Consulta()
		{
			if (this.cmbBodega.ActiveRow == null) return;

			string sSQL = string.Format("Exec [ConsultaConteoDiarioInventario] {0}, '{1}'", (int)this.cmbBodega.Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txidConteo.Value = dr.GetInt32(0);
				this.txtEstado.Value = dr.GetInt32(1);
				this.cmbVendedor.Value = dr.GetInt32(2);
				this.lblEstado.Text = dr.GetString(3);
				this.txtNotas.Text = dr.GetString(4);
				this.txtNumero.Text = dr.GetString(7);
			}
			dr.Close();	

			this.ListaConteo();
	
		}


		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			if ((int)this.txtEstado.Value == 2)
			{
				MessageBox.Show("El Conteo esta CERRADO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if ((int)this.txtEstado.Value == 3)
			{
				MessageBox.Show("El Conteo esta ANULADO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if ((int) this.txtEstado.Value == 0)
			{
				this.cmbBodega.Enabled = true;
				this.cmbSubGrupo.Enabled = true;
				this.cmbVendedor.Enabled = true;
				this.txtArticulo.Enabled = true;
				this.txtNotas.Enabled = true;
				this.btnGenerar.Enabled = true;
				this.btnVer.Enabled = true;

				FuncionesProcedimientos.EstadoGrids(true, this.uGridConteo);
			}
      
			if ((int) this.txtEstado.Value == 1)
			{
				this.cmbSubGrupo.Enabled = true;
				this.txtArticulo.Enabled = true;
				this.txtNotas.Enabled = true;
				this.btnVer.Enabled = true;
				FuncionesProcedimientos.EstadoGrids(true, this.uGridConteo);
			}


			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;

			string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones 0, 1, {0}, 3", (int)this.txidConteo.Value);
			idBloqueaTransacciones = Funcion.iEscalarSQL(cdsSeteoF, sSQLBloqueo, true);

			bNuevo = false;
			bEdicion = true;
		
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Cierre Conteo
			//			if (!Validacion.vbComboVacio(this.cmbVendedor, "Seleccione un vendedor")) return;

			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione la Bodega ")) return;
			if (!Validacion.vbComboVacio(this.cmbVendedor, "Seleccione el Responsable del Inventario")) return;
			if (!Validacion.vbTexto(this.txtNotas, 10, 100, "Notas")) return;

			if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de cerrar el conteo?\n\nEste proceso no puede ser revertido", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{							
				string sSQL = string.Format("Exec CierraConteoDiarioInventario {0}, '{1}', {2}",(int) this.txidConteo.Value, this.txtNotas.Text.ToString().Trim(), 0);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);									
							
				MessageBox.Show("Conteo de inventario cerrado correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				

				this.Consulta();
			}
			#endregion Cierre Conteo
			this.cmbBodega.Enabled = false;
			this.cmbSubGrupo.Enabled = false;
			this.cmbVendedor.Enabled = false;
			this.txtArticulo.Enabled = false;
			this.txtNotas.Enabled = false;
			this.btnGenerar.Enabled = false;
			this.btnVer.Enabled = false;
										
			FuncionesProcedimientos.EstadoGrids(false, this.uGridConteo);
										
			bNuevo = false;
			bEdicion = false;
		}

		private void cmbBodega_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (this.cmbBodega.ActiveRow == null)
				{
					MessageBox.Show("Seleccione la Bodega Origén", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.cmbBodega.Focus();
					e.Cancel = true;
				}
			}
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				if (this.cmbBodega.ActiveRow != null && this.dtFecha.Value != System.DBNull.Value)
				{
					bool bConsignacion = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select Consignacion From Bodega Where Bodega = {0}", (int)this.cmbBodega.Value), true);

					int iBodegaOrigen = (int)this.cmbBodega.Value;
					if (bConsignacion)
						iBodegaOrigen = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select Local from AsignaBodegasLocales Where Bodega = {0} And Consignacion = 1", (int)this.cmbBodega.Value));


					//			this.cmbVendedor.DataSource = FuncionesProcedimientos.Lista_Vendedores(Convert.ToDateTime(this.dtFecha.Value), (int)this.cmbBodega.Value, cdsSeteoF);

				
					if (iBodegaPredef == 0 && bNuevo && (MenuLatinium.iNivel != 23 && MenuLatinium.iNivel != 25))
						this.cmbVendedor.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("If Exists(Select idPersonal From Nomina Where Codigo In (Select Nombre From Usuario where idUsuario = {0})) Select idPersonal From Nomina Where Codigo In (Select Nombre From Usuario where idUsuario = {0}) Else Select 0", MenuLatinium.IdUsuario));					
					else
					{
						if (bEdicion) 
							this.cmbVendedor.Enabled = true;
					}
				}
			}
		}

		private void uGridConteo_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Exibicion")
			{				
				if (e.Cell.Row.Cells["Exibicion"].Value != System.DBNull.Value)
				{
					string sSQL = string.Format("Exec ActualizaConteoInventarioDiario {0}, {1}, '', 1", 
						(int)e.Cell.Row.Cells["idDetalleConteoDiarioInventario"].Value, (int)e.Cell.Row.Cells["Exibicion"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);

//					SumaConteo(e);
				}
			}

			if (e.Cell.Column.ToString() == "Bodega")
			{
				if (e.Cell.Row.Cells["Bodega"].Value != System.DBNull.Value)
				{
					string sSQL = string.Format("Exec ActualizaConteoInventarioDiario {0}, {1}, '', 2", 
						(int)e.Cell.Row.Cells["idDetalleConteoDiarioInventario"].Value, (int)e.Cell.Row.Cells["Bodega"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);

//					SumaConteo(e);
				}
			}

			if (e.Cell.Column.ToString() == "Conteo")
			{
				string sSQL = string.Format("Exec ActualizaConteoInventarioDiario {0}, {1}, '', 3", 
					(int)e.Cell.Row.Cells["idDetalleConteoDiarioInventario"].Value, (int)e.Cell.Row.Cells["Conteo"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);
			}

			if (e.Cell.Column.ToString() == "Notas")
			{
				string sSQL = string.Format("Exec ActualizaConteoInventarioDiario {0}, 0, '{1}', 4", 
					(int)e.Cell.Row.Cells["idDetalleConteoDiarioInventario"].Value, e.Cell.Row.Cells["Notas"].Value.ToString());
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);
			}
		}

		private void uGridConteo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridConteo);
		}

		private void frmConteoDiarioGreta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) this.btnNuevo_Click(sender, e);
//				if (e.KeyCode == Keys.B) this.btnConsultar_Click(sender, e);
				if ((int)this.txidConteo.Value > 0)
				{
					if (e.KeyCode == Keys.E) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);					
//					if (e.KeyCode == Keys.F5) this.Consultar((int)this.txidConteo.Value);					
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

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			this.ListaConteo();
		}

		private void txtArticulo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnVer_Click(sender, e);
		}

		private void cmbVendedor_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbSubGrupo.Focus();
		}

		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
						if (e.KeyCode == Keys.Enter) this.cmbSubGrupo.Focus();
		}

		private void cmbSubGrupo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
						if (e.KeyCode == Keys.Enter) this.txtNotas.Focus();
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			#region Imprime el reporte
			ParameterFields paramFields = new ParameterFields ();
				
			ParameterField pfRet = new ParameterField ();
			ParameterDiscreteValue id = new ParameterDiscreteValue ();
			pfRet.ParameterFieldName = "@idConteoDiarioInventario";
			id.Value = (int) this.txidConteo.Value;
			pfRet.CurrentValues.Add (id);
			paramFields.Add (pfRet);
				
			Reporte miRepor = new Reporte("ListaInventarioCDI.rpt", "");
			miRepor.MdiParent = this.MdiParent;			
			miRepor.crVista.ParameterFieldInfo = paramFields;			
			miRepor.Show();
			#endregion Imprime el reporte
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			using (frmConsultaInventarioGreta Busqueda = new frmConsultaInventarioGreta())
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.ConsultaR((int)Busqueda.uGrid.ActiveRow.Cells["idConteoDiarioInventario"].Value);
				}
			}

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = true;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = true;

		}

		private void ConsultaR (int Id)
		{			
			string sSQL = string.Format("Exec [ConsultaIndividualConteoDiarioInventario] {0}", Id);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txidConteo.Value = Id;
				if (dr.GetValue(0) != System.DBNull.Value) this.dtFecha.Value = dr.GetDateTime(0);
				if (dr.GetValue(1) != System.DBNull.Value) this.txtNumero.Text = dr.GetString(1); 
				if (dr.GetValue(3) != System.DBNull.Value) this.cmbBodega.Value = dr.GetInt32(3);
				if (dr.GetValue(4) != System.DBNull.Value) this.cmbVendedor.Value = dr.GetInt32(4);
				if (dr.GetValue(5) != System.DBNull.Value) this.txtEstado.Value = dr.GetInt32(5); 
				if (dr.GetValue(6) != System.DBNull.Value) this.txtNotas.Text = dr.GetString(6); 
			}
			dr.Close();

			if ((int)this.txtEstado.Value == 0 ) this.lblEstado.Text = "PENIDNETE";
			if ((int)this.txtEstado.Value == 1 ) this.lblEstado.Text = "ABIERTO";
			if ((int)this.txtEstado.Value == 2 ) this.lblEstado.Text = "CERRADO";

		}

	}
}
