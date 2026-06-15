using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using Infragistics.Win.UltraWinGrid;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmPlanificacionHorariosDeCarpa.
	/// </summary>
	public class frmPlanificacionHorariosDeCarpa : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCom_Ubicacion2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbVendedores;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridSemana;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstado;
		private System.Windows.Forms.Label lblContador;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotasAprobacion;
		private System.Windows.Forms.Button btnAprobar;
		private System.Windows.Forms.Button btnImprimir;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmPlanificacionHorariosDeCarpa()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmPlanificacionHorariosDeCarpa));
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleHorariosDeCarpa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idHorariosDeCarpa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Lunes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Martes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Miercoles");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Jueves");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Viernes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sabado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Domingo");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleHorariosDeCarpa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idHorariosDeCarpa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal", -1, "cmbVendedores");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lunes", -1, "cmbCom_Ubicacion2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Martes", -1, "cmbCom_Ubicacion2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Miercoles", -1, "cmbCom_Ubicacion2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jueves", -1, "cmbCom_Ubicacion2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Viernes", -1, "cmbCom_Ubicacion2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sabado", -1, "cmbCom_Ubicacion2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Domingo", -1, "cmbCom_Ubicacion2");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCom_Ubicacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ubicacion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Ubicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ubicacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idHorariosDeCarpa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Desde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Hasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoM");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NotasAprobacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Revisado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Imp");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Impreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NotasRevision");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idHorariosDeCarpa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desde");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hasta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoM");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NotasAprobacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Revisado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Imp");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NotasRevision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BtnRevision", 0);
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BtnImpreso", 1);
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			this.lblBodega = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridSemana = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbCom_Ubicacion2 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbVendedores = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblContador = new System.Windows.Forms.Label();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.txtNotasAprobacion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.optEstado = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btnAprobar = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSemana)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCom_Ubicacion2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedores)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotasAprobacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			this.SuspendLayout();
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(8, 10);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(31, 16);
			this.lblBodega.TabIndex = 649;
			this.lblBodega.Text = "Local";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 60;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 288;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(56, 8);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(220, 21);
			this.cmbBodega.TabIndex = 644;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// btnConsultar
			// 
			this.btnConsultar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnConsultar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(752, 7);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 761;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnEditar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(8, 272);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 763;
			this.btnEditar.Text = "Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(89, 272);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 762;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnCancelar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(248, 272);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 764;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(328, 272);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 765;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
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
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(int);
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
																																 ultraDataColumn10});
			// 
			// uGridSemana
			// 
			this.uGridSemana.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridSemana.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridSemana.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			appearance1.ForeColor = System.Drawing.Color.Black;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSemana.DisplayLayout.Appearance = appearance1;
			this.uGridSemana.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 116;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 89;
			ultraGridColumn6.Header.Caption = "Vendedor";
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn6.Width = 220;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn7.Header.VisiblePosition = 3;
			ultraGridColumn7.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn7.Width = 100;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn8.Header.VisiblePosition = 4;
			ultraGridColumn8.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn8.Width = 100;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn9.Header.VisiblePosition = 5;
			ultraGridColumn9.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn9.Width = 100;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn10.Header.VisiblePosition = 6;
			ultraGridColumn10.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn10.Width = 100;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn11.Header.VisiblePosition = 7;
			ultraGridColumn11.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn11.Width = 100;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn12.Header.VisiblePosition = 8;
			ultraGridColumn12.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn12.Width = 100;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn13.Header.VisiblePosition = 9;
			ultraGridColumn13.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn13.Width = 100;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13});
			ultraGridBand2.Header.Caption = "";
			this.uGridSemana.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSemana.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.uGridSemana.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridSemana.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridSemana.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.uGridSemana.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 7.5F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridSemana.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.uGridSemana.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance5.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance5.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemana.DisplayLayout.Override.RowAlternateAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemana.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemana.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.uGridSemana.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridSemana.Location = new System.Drawing.Point(8, 304);
			this.uGridSemana.Name = "uGridSemana";
			this.uGridSemana.Size = new System.Drawing.Size(941, 144);
			this.uGridSemana.TabIndex = 767;
			this.uGridSemana.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridSemana_AfterRowInsert);
			this.uGridSemana.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.uGridSemana_BeforeCellUpdate);
			this.uGridSemana.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridSemana_BeforeRowsDeleted);
			this.uGridSemana.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridSemana_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn11.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn11,
																																 ultraDataColumn12});
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
			// cmbCom_Ubicacion2
			// 
			this.cmbCom_Ubicacion2.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCom_Ubicacion2.DataSource = this.ultraDataSource2;
			ultraGridBand3.AddButtonCaption = "Proyecto";
			ultraGridColumn14.Header.VisiblePosition = 0;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 1;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn14,
																										 ultraGridColumn15});
			this.cmbCom_Ubicacion2.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbCom_Ubicacion2.DisplayMember = "Ubicacion";
			this.cmbCom_Ubicacion2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCom_Ubicacion2.Location = new System.Drawing.Point(536, 344);
			this.cmbCom_Ubicacion2.Name = "cmbCom_Ubicacion2";
			this.cmbCom_Ubicacion2.Size = new System.Drawing.Size(176, 86);
			this.cmbCom_Ubicacion2.TabIndex = 768;
			this.cmbCom_Ubicacion2.ValueMember = "idCom_Ubicacion";
			this.cmbCom_Ubicacion2.Visible = false;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn13.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15});
			// 
			// cmbVendedores
			// 
			this.cmbVendedores.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbVendedores.DataSource = this.ultraDataSource3;
			ultraGridColumn16.Header.VisiblePosition = 0;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn16.Width = 57;
			ultraGridColumn17.Header.VisiblePosition = 1;
			ultraGridColumn17.Width = 100;
			ultraGridColumn18.Header.VisiblePosition = 2;
			ultraGridColumn18.Width = 180;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			this.cmbVendedores.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbVendedores.DisplayMember = "Nombre";
			this.cmbVendedores.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbVendedores.Location = new System.Drawing.Point(176, 328);
			this.cmbVendedores.Name = "cmbVendedores";
			this.cmbVendedores.Size = new System.Drawing.Size(288, 112);
			this.cmbVendedores.TabIndex = 769;
			this.cmbVendedores.ValueMember = "idPersonal";
			this.cmbVendedores.Visible = false;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 40);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(960, 8);
			this.groupBox1.TabIndex = 770;
			this.groupBox1.TabStop = false;
			// 
			// txtNotas
			// 
			this.txtNotas.AcceptsReturn = true;
			this.txtNotas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance8;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Enabled = false;
			this.txtNotas.Location = new System.Drawing.Point(64, 456);
			this.txtNotas.MaxLength = 1000;
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(336, 66);
			this.txtNotas.TabIndex = 771;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.BackColor = System.Drawing.Color.Transparent;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.Location = new System.Drawing.Point(432, 274);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 19);
			this.lblContador.TabIndex = 772;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraDataSource4
			// 
			ultraDataColumn16.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn19.DataType = typeof(System.DateTime);
			ultraDataColumn20.DataType = typeof(System.DateTime);
			ultraDataColumn21.DataType = typeof(int);
			ultraDataColumn25.DataType = typeof(bool);
			ultraDataColumn26.DataType = typeof(bool);
			ultraDataColumn27.DataType = typeof(bool);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21,
																																 ultraDataColumn22,
																																 ultraDataColumn23,
																																 ultraDataColumn24,
																																 ultraDataColumn25,
																																 ultraDataColumn26,
																																 ultraDataColumn27,
																																 ultraDataColumn28});
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 480);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(34, 16);
			this.label2.TabIndex = 776;
			this.label2.Text = "Notas";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(432, 480);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(110, 16);
			this.label3.TabIndex = 778;
			this.label3.Text = "Notas de Aprobación";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtNotasAprobacion
			// 
			this.txtNotasAprobacion.AcceptsReturn = true;
			this.txtNotasAprobacion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotasAprobacion.Appearance = appearance9;
			this.txtNotasAprobacion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotasAprobacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotasAprobacion.Enabled = false;
			this.txtNotasAprobacion.Location = new System.Drawing.Point(552, 456);
			this.txtNotasAprobacion.MaxLength = 1000;
			this.txtNotasAprobacion.Multiline = true;
			this.txtNotasAprobacion.Name = "txtNotasAprobacion";
			this.txtNotasAprobacion.Size = new System.Drawing.Size(336, 66);
			this.txtNotasAprobacion.TabIndex = 777;
			// 
			// optEstado
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstado.Appearance = appearance10;
			this.optEstado.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstado.CheckedIndex = 0;
			this.optEstado.ItemAppearance = appearance11;
			this.optEstado.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Planificación Pendiente";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Aprobación Pendiente ";
			valueListItem4.DataValue = 3;
			valueListItem4.DisplayText = "Aprobado";
			this.optEstado.Items.Add(valueListItem1);
			this.optEstado.Items.Add(valueListItem2);
			this.optEstado.Items.Add(valueListItem3);
			this.optEstado.Items.Add(valueListItem4);
			this.optEstado.ItemSpacingVertical = 10;
			this.optEstado.Location = new System.Drawing.Point(296, 6);
			this.optEstado.Name = "optEstado";
			this.optEstado.Size = new System.Drawing.Size(440, 24);
			this.optEstado.TabIndex = 779;
			this.optEstado.Text = "Todos";
			this.optEstado.ValueChanged += new System.EventHandler(this.optEstado_ValueChanged);
			// 
			// btnAprobar
			// 
			this.btnAprobar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnAprobar.CausesValidation = false;
			this.btnAprobar.Enabled = false;
			this.btnAprobar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnAprobar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnAprobar.Image = ((System.Drawing.Image)(resources.GetObject("btnAprobar.Image")));
			this.btnAprobar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAprobar.Location = new System.Drawing.Point(168, 272);
			this.btnAprobar.Name = "btnAprobar";
			this.btnAprobar.TabIndex = 780;
			this.btnAprobar.Text = "Aprobar";
			this.btnAprobar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAprobar.Click += new System.EventHandler(this.btnAprobar_Click);
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnImprimir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(864, 272);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 781;
			this.btnImprimir.Text = "Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Visible = false;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource4;
			appearance12.BackColor = System.Drawing.Color.White;
			appearance12.ForeColor = System.Drawing.Color.Black;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Appearance = appearance12;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn19.Header.VisiblePosition = 0;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Width = 89;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn20.Header.VisiblePosition = 3;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn21.Header.VisiblePosition = 6;
			ultraGridColumn21.Width = 200;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn22.Header.VisiblePosition = 7;
			ultraGridColumn22.Width = 79;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn23.Header.VisiblePosition = 8;
			ultraGridColumn23.Width = 75;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn24.Header.VisiblePosition = 9;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn25.Header.Caption = "Estado";
			ultraGridColumn25.Header.VisiblePosition = 10;
			ultraGridColumn25.Width = 92;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn26.Header.Caption = "Notas";
			ultraGridColumn26.Header.VisiblePosition = 11;
			ultraGridColumn26.Width = 184;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn27.Header.Caption = "Notas de Aprobación";
			ultraGridColumn27.Header.VisiblePosition = 12;
			ultraGridColumn27.Width = 190;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn28.Header.VisiblePosition = 4;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn28.Width = 77;
			ultraGridColumn29.Header.Caption = "Sel-Imp.";
			ultraGridColumn29.Header.VisiblePosition = 5;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn29.Width = 70;
			ultraGridColumn30.Header.VisiblePosition = 14;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn30.Width = 53;
			ultraGridColumn31.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn31.Header.Caption = "Notas Revisión";
			ultraGridColumn31.Header.VisiblePosition = 13;
			ultraGridColumn31.Width = 190;
			ultraGridColumn32.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance13.Image = ((object)(resources.GetObject("appearance13.Image")));
			appearance13.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn32.CellButtonAppearance = appearance13;
			ultraGridColumn32.Header.Caption = "Revisado";
			ultraGridColumn32.Header.VisiblePosition = 1;
			ultraGridColumn32.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn32.Width = 45;
			ultraGridColumn33.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance14.Image = ((object)(resources.GetObject("appearance14.Image")));
			appearance14.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn33.CellButtonAppearance = appearance14;
			ultraGridColumn33.Header.Caption = "Sel-Impr.";
			ultraGridColumn33.Header.VisiblePosition = 2;
			ultraGridColumn33.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn33.Width = 45;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28,
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33});
			ultraGridBand5.Header.Caption = "";
			ultraGridBand5.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			appearance15.ForeColor = System.Drawing.Color.Black;
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance15;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance16.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance17.FontData.BoldAsString = "True";
			appearance17.FontData.Name = "Arial";
			appearance17.FontData.SizeInPoints = 7.5F;
			appearance17.ForeColor = System.Drawing.Color.White;
			appearance17.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance17;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance18.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance18.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance18;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance20;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 56);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(941, 208);
			this.ultraGrid1.TabIndex = 775;
			this.ultraGrid1.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.ultraGrid1_BeforeSelectChange);
			this.ultraGrid1.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_CellChange);
			this.ultraGrid1.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_ClickCellButton);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// frmPlanificacionHorariosDeCarpa
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(957, 534);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnAprobar);
			this.Controls.Add(this.optEstado);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.txtNotasAprobacion);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.cmbCom_Ubicacion2);
			this.Controls.Add(this.cmbVendedores);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.uGridSemana);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.Name = "frmPlanificacionHorariosDeCarpa";
			this.Text = "Planificación Semanal De Carpa";
			this.Load += new System.EventHandler(this.frmPlanificacionHorariosDeCarpa_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSemana)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCom_Ubicacion2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedores)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotasAprobacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}
		
		bool bEdicion = false;
		int iBodegaPredef = 0;
		int idHorariosDeCarpa = 0;
		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}
		bool bActualiza = false;
		private void frmPlanificacionHorariosDeCarpa_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0719");
			
			if (!Funcion.Permiso("649", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Planificación Semanal De Carpa", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));				

				return;
			}
			// acesso a revisado
			if (!miAcceso.CambiarNumeracion)
			{
				//				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Revisado"].CellActivation = Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["BtnRevision"].Hidden = true;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["NotasRevision"].Hidden = true;
				
			}
			else
			{
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["BtnRevision"].Hidden = false;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["NotasRevision"].Hidden = false;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Revisado"].Hidden = false;
				
			}
			this.cmbCom_Ubicacion2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idCom_Ubicacion, Ubicacion From Com_Ubicacion Where Activo = 1");


			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaBodegas {0}, 1", MenuLatinium.IdUsuario));
			iBodegaPredef = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario));

			if (iBodegaPredef == 0) this.cmbBodega.Enabled = true;
			else this.cmbBodega.Value = iBodegaPredef;

			FuncionesProcedimientos.EstadoGrids(false, this.uGridSemana);			
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.cmbBodega.Enabled = true;
			this.optEstado.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.ultraGrid1.Enabled = true;
			this.btnAprobar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnEditar.Enabled = false;

			idHorariosDeCarpa = 0;

			FuncionesProcedimientos.EstadoGrids(false, this.uGridSemana);

			this.txtNotasAprobacion.Enabled = false;
			this.txtNotas.Enabled = false;

			bEdicion = false;
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{			
			//Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update HorariosDeCarpa Set Estado = Case Estado When 3 Then Estado When 2 Then 3 When 1 Then 3 End, Observaciones = '{0}' Where idHorariosDeCarpa = {1}", this.txtNotas.Text.ToString(), idHorariosDeCarpa));

			int idPersonal = 0;
			
			int iLunes = 0;
			int iMartes = 0;
			int iMiercoles = 0;
			int iJueves = 0;
			int iViernes = 0;
			int iSabado = 0;
			int iDomingo = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridSemana.Rows.All)
			{
				iLunes = 0;
				iMartes = 0;
				iMiercoles = 0;
				iJueves = 0;
				iViernes = 0;
				iSabado = 0;
				iDomingo = 0;
				
				if (dr.Cells["idPersonal"].Value != System.DBNull.Value) idPersonal = (int)dr.Cells["idPersonal"].Value;
				if (dr.Cells["Lunes"].Value != System.DBNull.Value) iLunes = (int)dr.Cells["Lunes"].Value;
				if (dr.Cells["Martes"].Value != System.DBNull.Value) iMartes = (int)dr.Cells["Martes"].Value;
				if (dr.Cells["Miercoles"].Value != System.DBNull.Value) iMiercoles = (int)dr.Cells["Miercoles"].Value;
				if (dr.Cells["Jueves"].Value != System.DBNull.Value) iJueves = (int)dr.Cells["Jueves"].Value;
				if (dr.Cells["Viernes"].Value != System.DBNull.Value) iViernes = (int)dr.Cells["Viernes"].Value;
				if (dr.Cells["Sabado"].Value != System.DBNull.Value) iSabado = (int)dr.Cells["Sabado"].Value;
				if (dr.Cells["Domingo"].Value != System.DBNull.Value) iDomingo = (int)dr.Cells["Domingo"].Value;

				string sSQL = string.Format("Exec PlanificacionSemanalCarpaGuardaDetalle {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}",
					(int)dr.Cells["idDetalleHorariosDeCarpa"].Value, idHorariosDeCarpa, idPersonal, iLunes, iMartes, iMiercoles, iJueves, iViernes, iSabado, iDomingo);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);	
			}

			this.cmbBodega.Enabled = true;
			this.optEstado.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.ultraGrid1.Enabled = true;

			FuncionesProcedimientos.EstadoGrids(false, this.uGridSemana);

			this.txtNotasAprobacion.Enabled = false;
			this.txtNotas.Enabled = false;

			this.btnGuardar.Enabled = false;
			this.btnAprobar.Enabled = false;

			bEdicion = false;
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{			
			FuncionesProcedimientos.EstadoGrids(false, this.uGridSemana);
			
			if (!miAcceso.BAprobarCarpa) this.txtNotas.Enabled = true;
						
			this.cmbBodega.Enabled = false;
			this.optEstado.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.ultraGrid1.Enabled = false;
			this.txtNotas.Enabled = true;

			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;


			DateTime dtFechaServ = Funcion.dtEscalarSQL(cdsSeteoF, string.Format("Exec RetornaFechaServidor "));

//dtFechaServ Convert.ToDateTime(this.dtFechaSalida.Value).ToString("yyyyMMdd")

			int iDias = 0;
			 iDias =Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaDiaSemana '{0}'", dtFechaServ.ToString("yyyyMMdd")));

			FuncionesProcedimientos.EstadoGrids(true, this.uGridSemana);

			int iHora = DateTime.Now.Hour;

			if (iDias == 1) 
			{							
				if (iHora < 14)
					this.uGridSemana.DisplayLayout.Bands[0].Columns["Lunes"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;			
			}
			if (iDias <= 2)
			{
				if (iDias == 2)
				{
					if (iHora < 12)
						this.uGridSemana.DisplayLayout.Bands[0].Columns["Martes"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;								
				}
				else if (iDias < 2)
					this.uGridSemana.DisplayLayout.Bands[0].Columns["Martes"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;			
			}
			if (iDias <= 3) 
			{
				if (iDias == 3)
				{
					if (iHora < 12)
						this.uGridSemana.DisplayLayout.Bands[0].Columns["Miercoles"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;								
				}
				else if (iDias < 3)
					this.uGridSemana.DisplayLayout.Bands[0].Columns["Miercoles"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;			
			}
			if (iDias <= 4) 
			{
				if (iDias == 4)
				{
					if (iHora < 12)
						this.uGridSemana.DisplayLayout.Bands[0].Columns["Jueves"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;								
				}
				else if (iDias < 4)
					this.uGridSemana.DisplayLayout.Bands[0].Columns["Jueves"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;			
			}
			if (iDias <= 5) 
			{
				if (iDias == 5)
				{
					if (iHora < 12)
						this.uGridSemana.DisplayLayout.Bands[0].Columns["Viernes"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;								
				}
				else if (iDias < 5)
					this.uGridSemana.DisplayLayout.Bands[0].Columns["Viernes"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;			
			}
			if (iDias <= 6)
			{
				if (iDias == 6)
				{
					if (iHora < 12)
						this.uGridSemana.DisplayLayout.Bands[0].Columns["Sabado"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;								
				}
				else if (iDias < 6)
					this.uGridSemana.DisplayLayout.Bands[0].Columns["Sabado"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;			
			}				
			if (iDias <= 7)
			{
				if (iDias == 7)
				{
					if (iHora < 12)
						this.uGridSemana.DisplayLayout.Bands[0].Columns["Domingo"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;								
				}
				else if (iDias < 7)
					this.uGridSemana.DisplayLayout.Bands[0].Columns["Domingo"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;			
			}

			bEdicion = true;

			//			if (miAcceso.BAprobarCarpa) this.txtNotasAprobacion.Enabled = true;
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{				
			Consulta();
		}

		private void ultraGrid1_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Revisado 
			if (e.Cell.Column.ToString() == "BtnRevision")
			{	
				using (frmNotas Notas  = new frmNotas())
				{
					if (DialogResult.OK == Notas.ShowDialog())
					{
						e.Cell.Row.Cells["NotasRevision"].Value = Notas.txtNotas.Text;

						string sSQL = string.Format("Update HorariosDeCarpa  set NotasRevision = '{0}', Revisado = 1  where idHorariosDeCarpa = {1}", 
							Convert.ToString ( e.Cell.Row.Cells["NotasRevision"].Value),(int)e.Cell.Row.Cells["idHorariosDeCarpa"].Value );
						Funcion.EjecutaSQL(cdsSeteoF, sSQL);

						this.btnConsultar_Click (sender, e);
					}
				}
			}
			#endregion Revisado

			#region Imprimir 

			if (e.Cell.Column.ToString() == "BtnImpreso")
			{	

				if ((int)e.Cell.Row.Cells["Estado"].Value < 3 )
				{
					MessageBox.Show("Comuniquese con su Supervisor de ventas para su respectiva APROBACIÓN ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//					e.Cell.Row.Cells["Imp"].Value = (bool)e.Cell.OriginalValue;
					return;
				}
				if (  (bool)e.Cell.Row.Cells["Impreso"].Value == true  )
				{
					if (!miAcceso.BExportar) 
					{
						MessageBox.Show("La Planificación seleccionada ya fue Impresa ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
													
						return;
					}
				}

				if (DialogResult.Yes == MessageBox.Show(" ¿Desea imprimir la Planificacion Semanal Seleccionada?, RECUERDE QUE LA IMPRESIÓN SALDRA DIRECTAMENTE  A LA IMPRESORA PREDETERMINADA ,  Si presiona SI ya no podra volver a Reimprimir", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
				{
					#region Impresion Pantalla
//					ParameterFields paramFields = new ParameterFields();
//
//					//			ParameterFields paramFields = new ParameterFields ();
//					ParameterField paramFieldCedula = new ParameterField ();
//					ParameterDiscreteValue ValCedula = new ParameterDiscreteValue ();
//					paramFieldCedula.ParameterFieldName = "@idHorariosCarpa";
//					ValCedula.Value = (int) e.Cell.Row.Cells["idHorariosDeCarpa"].Value;
//					paramFieldCedula.CurrentValues.Add (ValCedula);
//					paramFields.Add (paramFieldCedula);
//
//					Reporte miRep = new Reporte("PlanificacionSemanalCarpa.rpt", "");
//					miRep.MdiParent = this.MdiParent;
//					miRep.Titulo("Planificacion De Carpa");
//					miRep.crVista.ParameterFieldInfo = paramFields;
//					miRep.Show();
					#endregion Impresion Pantalla

					#region Impresion
					string sReporte = "PlanificacionSemanalCarpa.rpt";
					ReportDocument Report1 = new ReportDocument();
					Report1.FileName = MenuLatinium.stDirReportes + sReporte;//"RolesIndividuales.rpt";
								
					ConnectionInfo crConnectionInfo = new ConnectionInfo (); 
				{
					string strBaseDato = MenuLatinium.stDirNumero;
					string strUsuario = MenuLatinium.stDirConta;
					string strServidor = MenuLatinium.stDirServidor;
					string strImpresora = MenuLatinium.ImpresoraNombre;

					crConnectionInfo.ServerName = strServidor;
					crConnectionInfo.DatabaseName = strBaseDato; 
					crConnectionInfo.UserID = "infoelect"; 
					crConnectionInfo.Password = "Bl45o6$9"; 
				}
			
					TableLogOnInfo crTableLogOnInfo; 
					Database crDatabase = Report1.Database; 
					Tables crTables = crDatabase.Tables; 
					Table crTable;

					for (int i=0; i< crTables.Count; i++)
					{
					{
						crTable = crTables [i]; 
						crTableLogOnInfo = crTable.LogOnInfo; 
						crTableLogOnInfo.ConnectionInfo = crConnectionInfo; 
						crTable.ApplyLogOnInfo(crTableLogOnInfo);
					}
					}
					
					Report1.PrintOptions.PrinterName = MenuLatinium.ImpresoraNombre;
					Report1.SetParameterValue("@idHorariosCarpa", (int) e.Cell.Row.Cells["idHorariosDeCarpa"].Value);
					
					Report1.PrintToPrinter(1, false, 0, 0);
			
					#endregion Impresion

					Funcion.EjecutaSQL(cdsSeteoF, string.Format("EXEC [GuardaImpresionPlanificacionSemanalCarpa] {0} ", (int) e.Cell.Row.Cells["idHorariosDeCarpa"].Value));			
			  
					this.btnConsultar_Click (sender, e);
				}
				
			}
			#endregion Imprimir
		}

		private void Consulta()
		{			
			int iBodega = 0;
			if (this.cmbBodega.ActiveRow != null) iBodega = (int)this.cmbBodega.Value;

			string sSQL = string.Format("Exec PlanificacionSemanalCarpaConsulta {0}, {1}", iBodega, (int)this.optEstado.Value);
			FuncionesProcedimientos.dsGeneral(sSQL, this.ultraGrid1);

			this.lblContador.Text = this.ultraGrid1.Rows.Count + " Registros Encontrados";
		}

		private void optEstado_ValueChanged(object sender, System.EventArgs e)
		{
			Consulta();

			FuncionesProcedimientos.dsGeneral(string.Format("Exec PlanificacionSemanalCarpaConsultaDetalle {0}", 0), this.uGridSemana);

			if ((int)this.optEstado.Value == 2)
			{
				if (miAcceso.BAprobarCarpa) 
				{
					this.txtNotasAprobacion.Enabled = true;		
					this.btnAprobar.Enabled = true;
				}
			}
			else
			{
				this.txtNotasAprobacion.Enabled = false;
				this.btnAprobar.Enabled = false;
			}
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			Consulta();

			FuncionesProcedimientos.dsGeneral(string.Format("Exec PlanificacionSemanalCarpaConsultaDetalle {0}", 0), this.uGridSemana);
		}

		private void ultraGrid1_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
			this.btnCancelar.Enabled = true;

			this.cmbVendedores.DataSource = FuncionesProcedimientos.Lista_VendedoresPlanifCarpa(Convert.ToDateTime(this.ultraGrid1.ActiveRow.Cells["Desde"].Value), (int)this.ultraGrid1.ActiveRow.Cells["Bodega"].Value, cdsSeteoF);

			idHorariosDeCarpa = (int)this.ultraGrid1.ActiveRow.Cells["idHorariosDeCarpa"].Value;

			FuncionesProcedimientos.dsGeneral(string.Format("Exec PlanificacionSemanalCarpaConsultaDetalle {0}", idHorariosDeCarpa), this.uGridSemana);
			
			string sSQL = string.Format("Select COUNT(1) From dbo.HorariosDeCarpa Where idHorariosDeCarpa = {0} And '{1}' Between CONVERT(Date, Desde) And CONVERT(Date, Hasta)", 
				(int)this.ultraGrid1.ActiveRow.Cells["idHorariosDeCarpa"].Value, DateTime.Today.ToString("yyyyMMdd"));

			if (Funcion.iEscalarSQL(cdsSeteoF, sSQL, true) > 0)
			{
				if ((int)this.ultraGrid1.ActiveRow.Cells["Estado"].Value == 1) this.btnEditar.Enabled = true;
				if ((int)this.ultraGrid1.ActiveRow.Cells["Estado"].Value == 2 && miAcceso.BAprobarCarpa) this.btnEditar.Enabled = true;
				if ((int)this.ultraGrid1.ActiveRow.Cells["Estado"].Value == 3 && miAcceso.BAprobarCarpa) this.btnEditar.Enabled = true;
			}

			if (miAcceso.BAprobarCarpa) this.btnAprobar.Enabled = true;
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridSemana_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btnAprobar_Click(object sender, System.EventArgs e)
		{
			if (this.ultraGrid1.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un registro de Carpa para Aprobar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update HorariosDeCarpa Set Estado = 3, NotasAprobacion = '{0}', FechaAprobacion = GETDATE(), UsuarioAprobo = '{2}' Where idHorariosDeCarpa = {1}", 
				this.txtNotasAprobacion.Text.ToString(), idHorariosDeCarpa, MenuLatinium.stUsuario));

			this.btnAprobar.Enabled = false;

			MessageBox.Show("Registro de Carpa Aprobado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

			Consulta();
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridSemana_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idDetalleHorariosDeCarpa"].Value = 0;
			e.Row.Cells["idHorariosDeCarpa"].Value = 0;
			e.Row.Cells["idPersonal"].Value = 0;			
			e.Row.Cells["Lunes"].Value = 0;
			e.Row.Cells["Martes"].Value = 0;
			e.Row.Cells["Miercoles"].Value = 0;
			e.Row.Cells["Jueves"].Value = 0;
			e.Row.Cells["Viernes"].Value = 0;
			e.Row.Cells["Sabado"].Value = 0;
			e.Row.Cells["Domingo"].Value = 0;			
		}

		private void uGridSemana_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			e.DisplayPromptMsg = false;
			if (DialogResult.Yes ==	MessageBox.Show("¿Esta Seguro de Borrar este registro?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				for (int i=0; i<e.Rows.Length; i++)
				{
					string sSQL = string.Format("Delete From DetalleHorariosDeCarpa Where idDetalleHorariosDeCarpa = {0}", (int)e.Rows[i].Cells["idDetalleHorariosDeCarpa"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
				}
			}
			else e.Cancel = true;
		}

		private void uGridSemana_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			if (e.Cell.Column.ToString() == "idPersonal")
			{
				if (e.Cell.Row.Cells["idPersonal"].Value == System.DBNull.Value) return;
	
				int iDias = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaDiaSemana '{0}'", DateTime.Today.ToString("yyyyMMdd")));
				int iHora = DateTime.Now.Hour;
			
				if (iDias == 1)
					if ((int)e.Cell.Row.Cells["idPersonal"].Value > 0)	
						if (iHora > 11) e.Cancel = true;

				if (iDias > 1)
					if ((int)e.Cell.Row.Cells["idPersonal"].Value > 0) e.Cancel = true;
			}
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bActualiza) return;
			if (e.Cell.Column.ToString() == "Revisado")
			{		
				bActualiza = true;
				if((bool)	this.ultraGrid1.ActiveRow.Cells["Revisado"].Value  ==  true )
				{
					string sSQL = string.Format("UPDATE HorariosDeCarpa SET Revisado = 1, FechaRevision = GETDATE(), UsuarioReviso = '{1}' WHERE idHorariosDeCarpa = {0}", 
						(int)	this.ultraGrid1.ActiveRow.Cells["idHorariosDeCarpa"].Value, MenuLatinium.stUsuario);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);	
				}
				else
				{
					e.Cell.Row.Cells["Revisado"].Value = (bool)e.Cell.OriginalValue;
					
				}
				bActualiza = false;
			}
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			//			int cont = 0;
			//			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			//			{
			//				if ((bool) dr.Cells["Imp"].Value)
			//				{
			//					cont = cont + 1;
			//				}
			//			}
			//			if (cont <1)
			//			{
			//				MessageBox.Show("Seleccione al menos una Fecha ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//				return;
			//			}
			//
			//			
			//			#region Valida Local
			//
			//			string sLocal= "";
			//			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			//			{
			//				if ((bool) dr.Cells["Imp"].Value)
			//				{
			//				
			//					ParameterFields paramFields = new ParameterFields();
			//
			//					//			ParameterFields paramFields = new ParameterFields ();
			//				
			//					ParameterField paramFieldCedula = new ParameterField ();
			//					ParameterDiscreteValue ValCedula = new ParameterDiscreteValue ();
			//					paramFieldCedula.ParameterFieldName = "@idHorariosCarpa";
			//					ValCedula.Value = (int) dr.Cells["idHorariosDeCarpa"].Value;
			//					paramFieldCedula.CurrentValues.Add (ValCedula);
			//					paramFields.Add (paramFieldCedula);
			//
			//					Reporte miRep = new Reporte("PlanificacionSemanalCarpa.rpt", "");
			//					miRep.MdiParent = this.MdiParent;
			//					miRep.Titulo("Planificacion De Carpa");
			//					miRep.crVista.ParameterFieldInfo = paramFields;
			//					miRep.Show();
			//
			//					Funcion.EjecutaSQL(cdsSeteoF, string.Format("EXEC [GuardaImpresionPlanificacionSemanalCarpa] {0} ", (int) dr.Cells["idHorariosDeCarpa"].Value));			
			//				}
			//		}
			//
			//			#endregion Valida Local
			//
			//			this.btnConsultar_Click (sender, e);


		}

		private void ultraGrid1_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			//			if (e.Cell.Column.ToString() == "Imp")
			//			{	
			//				this.ultraGrid1.UpdateData();
			//							
			//				if ((int)e.Cell.Row.Cells["Estado"].Value < 3 )
			//				{
			//					MessageBox.Show("Comuniquese con su Supervisor de ventas para su respectiva APROBACIÓN ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//					e.Cell.Row.Cells["Imp"].Value = (bool)e.Cell.OriginalValue;
			//				}
			//				if ( (bool)e.Cell.Row.Cells["Impreso"].Value == true)
			//				{
			//					MessageBox.Show("La Planificación seleccionada ya fue Impresa ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//								
			//					e.Cell.Row.Cells["Imp"].Value = (bool)e.Cell.OriginalValue;
			//				}
			//			}
		}

		
	}
}
