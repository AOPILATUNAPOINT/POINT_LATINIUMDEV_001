using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frm_consulta_planificacion_pagos.
	/// </summary>
	public class frm_consulta_planificacion_pagos : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnSalir;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbProyecto;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbIdCuenta;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCuentaContable;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridPagos;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbTipos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbOrigenAsientos;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbPeriodos;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frm_consulta_planificacion_pagos()
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
 	private Acceso miAcceso;
		#region Código generado por el Diseñador de Windows Forms
		/// <summary>
		/// Método necesario para admitir el Diseñador. No se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frm_consulta_planificacion_pagos));
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPlanificacionPagos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoPlanificacionPagos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuentaBancaria");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrigenAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuentaContable");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPlanificacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Concepto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DiaEmision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DiaCancelacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VMA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Activo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaSistema");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdPlanificacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdPlanificacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPlanificacionPagos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoPlanificacionPagos", -1, "cmbTipos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaBancaria", -1, "cmbIdCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrigenAsiento", -1, "cmbOrigenAsientos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaContable", -1, "cmbCuentaContable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto", -1, "cmbProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPlanificacion", -1, "cmbPeriodos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Concepto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiaEmision");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiaCancelacion");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VMA");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Activo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSistema");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoPlanificacionPagos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoPlanificacionPagos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrigenAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrigenAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbProyecto = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbIdCuenta = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbCuentaContable = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbPeriodos = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.uGridPagos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cmbTipos = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbOrigenAsientos = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbIdCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaContable)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPeriodos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPagos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbOrigenAsientos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			this.SuspendLayout();
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
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGuardar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(998, 8);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 765;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
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
			this.btnSalir.Location = new System.Drawing.Point(1078, 8);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 766;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(bool);
			ultraDataColumn13.DefaultValue = false;
			ultraDataColumn14.DataType = typeof(System.Decimal);
			ultraDataColumn15.DataType = typeof(bool);
			ultraDataColumn15.DefaultValue = false;
			ultraDataColumn17.DataType = typeof(System.DateTime);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19});
			// 
			// cmbProyecto
			// 
			this.cmbProyecto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProyecto.DataSource = this.ultraDataSource3;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 200;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbProyecto.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbProyecto.DisplayMember = "CentroCosto";
			this.cmbProyecto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbProyecto.Location = new System.Drawing.Point(360, 112);
			this.cmbProyecto.Name = "cmbProyecto";
			this.cmbProyecto.Size = new System.Drawing.Size(136, 112);
			this.cmbProyecto.TabIndex = 774;
			this.cmbProyecto.ValueMember = "IdProyecto";
			this.cmbProyecto.Visible = false;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn20.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn20,
																																 ultraDataColumn21});
			// 
			// cmbIdCuenta
			// 
			this.cmbIdCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbIdCuenta.DataSource = this.ultraDataSource1;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 250;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbIdCuenta.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbIdCuenta.DisplayMember = "Descripcion";
			this.cmbIdCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbIdCuenta.Location = new System.Drawing.Point(568, 112);
			this.cmbIdCuenta.Name = "cmbIdCuenta";
			this.cmbIdCuenta.Size = new System.Drawing.Size(104, 112);
			this.cmbIdCuenta.TabIndex = 775;
			this.cmbIdCuenta.ValueMember = "IdCuenta";
			this.cmbIdCuenta.Visible = false;
			// 
			// cmbCuentaContable
			// 
			this.cmbCuentaContable.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentaContable.DataSource = this.ultraDataSource4;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 250;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbCuentaContable.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbCuentaContable.DisplayMember = "Descripcion";
			this.cmbCuentaContable.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbCuentaContable.Location = new System.Drawing.Point(176, 120);
			this.cmbCuentaContable.Name = "cmbCuentaContable";
			this.cmbCuentaContable.Size = new System.Drawing.Size(104, 112);
			this.cmbCuentaContable.TabIndex = 779;
			this.cmbCuentaContable.ValueMember = "IdCuenta";
			this.cmbCuentaContable.Visible = false;
			// 
			// ultraDataSource4
			// 
			ultraDataColumn22.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn22,
																																 ultraDataColumn23});
			// 
			// cmbPeriodos
			// 
			this.cmbPeriodos.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPeriodos.DataSource = this.ultraDataSource5;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Width = 150;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.cmbPeriodos.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbPeriodos.DisplayMember = "Descripcion";
			this.cmbPeriodos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbPeriodos.Location = new System.Drawing.Point(688, 160);
			this.cmbPeriodos.Name = "cmbPeriodos";
			this.cmbPeriodos.Size = new System.Drawing.Size(136, 112);
			this.cmbPeriodos.TabIndex = 783;
			this.cmbPeriodos.ValueMember = "IdPlanificacion";
			this.cmbPeriodos.Visible = false;
			// 
			// ultraDataSource5
			// 
			ultraDataColumn24.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn24,
																																 ultraDataColumn25});
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 40);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1158, 8);
			this.groupBox1.TabIndex = 785;
			this.groupBox1.TabStop = false;
			// 
			// uGridPagos
			// 
			this.uGridPagos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridPagos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridPagos.DataSource = this.ultraDataSource2;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridPagos.DisplayLayout.Appearance = appearance1;
			this.uGridPagos.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 125;
			ultraGridColumn10.Header.Caption = "Tipo";
			ultraGridColumn10.Header.VisiblePosition = 2;
			ultraGridColumn10.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn10.Width = 120;
			ultraGridColumn11.Header.Caption = "Cuenta Bancaria";
			ultraGridColumn11.Header.VisiblePosition = 3;
			ultraGridColumn11.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn11.Width = 200;
			ultraGridColumn12.Header.Caption = "Asiento Referencial";
			ultraGridColumn12.Header.VisiblePosition = 4;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn12.Width = 220;
			ultraGridColumn13.Header.Caption = "Cuenta Contable";
			ultraGridColumn13.Header.VisiblePosition = 5;
			ultraGridColumn13.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn13.Width = 200;
			ultraGridColumn14.Header.Caption = "Centro De Costo";
			ultraGridColumn14.Header.VisiblePosition = 6;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn14.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn14.Width = 200;
			ultraGridColumn15.Header.Caption = "Frecuencia";
			ultraGridColumn15.Header.VisiblePosition = 7;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn15.Width = 100;
			ultraGridColumn16.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn16.Header.VisiblePosition = 10;
			ultraGridColumn16.MaskInput = "";
			ultraGridColumn16.Width = 300;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn17.CellAppearance = appearance2;
			ultraGridColumn17.Header.Caption = "Desde";
			ultraGridColumn17.Header.VisiblePosition = 8;
			ultraGridColumn17.Width = 75;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn18.CellAppearance = appearance3;
			ultraGridColumn18.Header.Caption = "Hasta";
			ultraGridColumn18.Header.VisiblePosition = 9;
			ultraGridColumn18.Width = 75;
			ultraGridColumn19.Header.Caption = "V M A";
			ultraGridColumn19.Header.VisiblePosition = 11;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance4;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.VisiblePosition = 12;
			ultraGridColumn20.Width = 80;
			ultraGridColumn21.Header.VisiblePosition = 1;
			ultraGridColumn21.Width = 25;
			ultraGridColumn22.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn22.Header.VisiblePosition = 13;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn22.Width = 300;
			ultraGridColumn23.Header.VisiblePosition = 14;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 15;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 16;
			ultraGridColumn25.Hidden = true;
			ultraGridBand5.Columns.AddRange(new object[] {
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
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25});
			this.uGridPagos.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.uGridPagos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridPagos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridPagos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance5.BackColor = System.Drawing.Color.Transparent;
			this.uGridPagos.DisplayLayout.Override.CardAreaAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 8.5F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridPagos.DisplayLayout.Override.HeaderAppearance = appearance6;
			this.uGridPagos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPagos.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPagos.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.uGridPagos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridPagos.Location = new System.Drawing.Point(8, 56);
			this.uGridPagos.Name = "uGridPagos";
			this.uGridPagos.Size = new System.Drawing.Size(1142, 248);
			this.uGridPagos.TabIndex = 788;
			this.uGridPagos.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridPagos_CellChange);
			this.uGridPagos.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridPagos_AfterRowInsert);
			this.uGridPagos.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.uGridPagos_KeyPress);
			this.uGridPagos.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridPagos_BeforeRowsDeleted);
			this.uGridPagos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridPagos_KeyDown);
			this.uGridPagos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridPagos_AfterCellUpdate);
			this.uGridPagos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridPagos_InitializeLayout);
			// 
			// cmbTipos
			// 
			this.cmbTipos.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipos.DataSource = this.ultraDataSource6;
			ultraGridColumn26.Header.VisiblePosition = 0;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 1;
			ultraGridColumn27.Width = 180;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn26,
																										 ultraGridColumn27});
			this.cmbTipos.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbTipos.DisplayMember = "Tipo";
			this.cmbTipos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbTipos.Location = new System.Drawing.Point(32, 112);
			this.cmbTipos.MaxDropDownItems = 20;
			this.cmbTipos.Name = "cmbTipos";
			this.cmbTipos.Size = new System.Drawing.Size(128, 112);
			this.cmbTipos.TabIndex = 789;
			this.cmbTipos.ValueMember = "idTipoPlanificacionPagos";
			this.cmbTipos.Visible = false;
			// 
			// ultraDataSource6
			// 
			ultraDataColumn26.DataType = typeof(int);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn26,
																																 ultraDataColumn27});
			// 
			// cmbOrigenAsientos
			// 
			this.cmbOrigenAsientos.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbOrigenAsientos.DataSource = this.ultraDataSource7;
			ultraGridColumn28.Header.VisiblePosition = 0;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 1;
			ultraGridColumn29.Width = 250;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn28,
																										 ultraGridColumn29});
			this.cmbOrigenAsientos.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbOrigenAsientos.DisplayMember = "Descripcion";
			this.cmbOrigenAsientos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbOrigenAsientos.Location = new System.Drawing.Point(880, 128);
			this.cmbOrigenAsientos.Name = "cmbOrigenAsientos";
			this.cmbOrigenAsientos.Size = new System.Drawing.Size(104, 112);
			this.cmbOrigenAsientos.TabIndex = 790;
			this.cmbOrigenAsientos.ValueMember = "idOrigenAsiento";
			this.cmbOrigenAsientos.Visible = false;
			// 
			// ultraDataSource7
			// 
			ultraDataColumn28.DataType = typeof(int);
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn28,
																																 ultraDataColumn29});
			// 
			// frm_consulta_planificacion_pagos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1158, 310);
			this.Controls.Add(this.cmbOrigenAsientos);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.cmbPeriodos);
			this.Controls.Add(this.cmbIdCuenta);
			this.Controls.Add(this.cmbProyecto);
			this.Controls.Add(this.cmbCuentaContable);
			this.Controls.Add(this.cmbTipos);
			this.Controls.Add(this.uGridPagos);
			this.KeyPreview = true;
			this.Name = "frm_consulta_planificacion_pagos";
			this.Text = "Planificación Mensual de Pagos";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frm_consulta_planificacion_pagos_KeyDown);
			this.Load += new System.EventHandler(this.frm_consulta_planificacion_pagos_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbIdCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaContable)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPeriodos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPagos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbOrigenAsientos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		
		private void UnloadMe()
		{
			this.Close();
		}
		private void Consulta()
		{
			this.uGridPagos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaPlanificacionPagos");
		}
		
		private void frm_consulta_planificacion_pagos_Load(object sender, System.EventArgs e)
		{		
			miAcceso = new Acceso(cdsSeteoF, "044401");
			
			if (!Funcion.Permiso("914", cdsSeteoF))
			{				
				
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			this.cmbIdCuenta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select c.idCuenta, Descripcion From Cuenta c Inner Join CuentaCorriente cc On cc.idCuenta = c.idCuenta Where cc.Flujo = 1 Order By Descripcion");

			this.cmbCuentaContable.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idcuenta, Descripcion From Cuenta Where Grupo = 0 and Tipo In (2, 5) Order By Codigo");

			//this.cmbOrigenAsientos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "SELECT idOrigenAsiento, Descripcion FROM OrigenAsientos ORDER BY Descripcion");

			this.cmbProyecto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListaCentroDeCostos");

			this.cmbTipos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idTipoPlanificacionPagos, Tipo From TipoPlanificacionPagos Order By Tipo");

			this.cmbPeriodos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "ConsultaT");

			Consulta();
		}		

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if (this.uGridPagos.Rows.Count == 0)
			{
				MessageBox.Show("No existen registros para guardar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
				return;
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridPagos.Rows.All)
			{
				if (dr.Cells["idTipoPlanificacionPagos"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione un tipo de pago", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
					this.uGridPagos.ActiveRow = dr;
					this.uGridPagos.ActiveRow.Selected = true;					
					return;
				}
				if (dr.Cells["idCuentaBancaria"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione una Cuenta Bancaria", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
					this.uGridPagos.ActiveRow = dr;
					this.uGridPagos.ActiveRow.Selected = true;					
					return;
				}
				if (dr.Cells["idOrigenAsiento"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione un Asiento Referencial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
					this.uGridPagos.ActiveRow = dr;
					this.uGridPagos.ActiveRow.Selected = true;					
					return;
				}
//				if (dr.Cells["idCuentaContable"].Value == System.DBNull.Value)
//				{
//					MessageBox.Show("Seleccione una Cuenta Contable", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
//					this.uGridPagos.ActiveRow = dr;
//					this.uGridPagos.ActiveRow.Selected = true;					
//					return;
//				}
//				if (dr.Cells["idProyecto"].Value == System.DBNull.Value)
//				{
//					MessageBox.Show("Seleccione un Centro De Costo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
//					this.uGridPagos.ActiveRow = dr;
//					this.uGridPagos.ActiveRow.Selected = true;					
//					return;
//				}
//				if (dr.Cells["idPlanificacion"].Value == System.DBNull.Value)
//				{
//					MessageBox.Show("Seleccione la frecuencia de reserva", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//					this.uGridPagos.ActiveRow = dr;
//					this.uGridPagos.ActiveRow.Selected = true;					
//					return;
//				}
				if (dr.Cells["DiaEmision"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Debe ingresar el día de emisión del pago", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
					this.uGridPagos.ActiveRow = dr;
					this.uGridPagos.ActiveRow.Selected = true;
					return;
				}
				if ((int)dr.Cells["DiaEmision"].Value > 30)
				{
					MessageBox.Show("El día de emisión del pago no puede ser mayor a 30", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
					this.uGridPagos.ActiveRow = dr;
					this.uGridPagos.ActiveRow.Selected = true;
					return;
				}
				if ((int)dr.Cells["DiaEmision"].Value > (int)dr.Cells["DiaCancelacion"].Value)
				{
					MessageBox.Show("El día de emisión del pago no puede ser mayor al día de cancelación del pago", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
					this.uGridPagos.ActiveRow = dr;
					this.uGridPagos.ActiveRow.Selected = true;
					return;
				}
				if (dr.Cells["DiaCancelacion"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Debe ingresar el día de cancelación del pago", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
					this.uGridPagos.ActiveRow = dr;
					this.uGridPagos.ActiveRow.Selected = true;
					return;
				}
				if ((int)dr.Cells["DiaCancelacion"].Value > 30)
				{
					MessageBox.Show("El día de cancelación del pago no puede ser mayor a 30", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
					this.uGridPagos.ActiveRow = dr;
					this.uGridPagos.ActiveRow.Selected = true;
					return;
				}				
				if (dr.Cells["Concepto"].Value.ToString().Length == 0)
				{
					MessageBox.Show("Debe ingresar un concepto", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						
					this.uGridPagos.ActiveRow = dr;
					this.uGridPagos.ActiveRow.Selected = true;
					return;
				}
				if (!(bool)dr.Cells["VMA"].Value)
				{
					if (Convert.ToDecimal(dr.Cells["Valor"].Value) == 0.00m)
					{
						MessageBox.Show("El valor del pago no puede ser 0", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridPagos.ActiveRow = dr;
						this.uGridPagos.ActiveRow.Selected = true;					
						return;
					}
				}				
			}
			#endregion Validacion

			#region Guarda Registros
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridPagos.Rows.All)
			{
				int idCuenta = 0;
				if (dr.Cells["idCuentaContable"].Value != System.DBNull.Value) idCuenta = (int)dr.Cells["idCuentaContable"].Value;

				string sSQLDetalle = string.Format("Exec GuardaPlanificacionPagos {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', {8}, {9}, {10}, {11}, {12}, '{13}'",
					(int)dr.Cells["idPlanificacionPagos"].Value, (int)dr.Cells["idTipoPlanificacionPagos"].Value,
					(int)dr.Cells["idCuentaBancaria"].Value, (int)dr.Cells["idOrigenAsiento"].Value, idCuenta,
					(int)dr.Cells["idProyecto"].Value, (int)dr.Cells["idPlanificacion"].Value, 
					dr.Cells["Concepto"].Value.ToString(), (int)dr.Cells["DiaEmision"].Value, (int)dr.Cells["DiaCancelacion"].Value, 
					(bool)dr.Cells["VMA"].Value, (decimal)dr.Cells["Valor"].Value, (bool)dr.Cells["Activo"].Value, 
					dr.Cells["Observacion"].Value.ToString());
				Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalle);
				/*
				 @idPlanificacionPagos Int, @idTipoPlanificacionPagos Int, @idCuentaBancaria Int, @idCuentaContable Int, @idProyecto int, @idPlanificacion int,
				@Concepto Varchar(100), @DiaEmision Int, @DiaCancelacion Int, @VMA Bit, @Valor Decimal (18, 2), @Activo Bit, @Observacion Varchar(300)
				*/
			}
			#endregion Guarda Registros

			MessageBox.Show("Registro Guardado Correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

			Consulta();
		}

		private void uGridPagos_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idPlanificacionPagos"].Value = 0;
			e.Row.Cells["idOrigenAsiento"].Value = 0;
			e.Row.Cells["idCuentaContable"].Value = 0;
			e.Row.Cells["idProyecto"].Value = 0;
			e.Row.Cells["idPlanificacion"].Value = 0;			
			e.Row.Cells["DiaEmision"].Value = 1;
			e.Row.Cells["DiaCancelacion"].Value = 1;
			e.Row.Cells["Valor"].Value = 0.00m;
			e.Row.Cells["VMA"].Value = false;
			e.Row.Cells["Activo"].Value = true;
		}

		private void uGridPagos_KeyDown_1(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender,e,this.uGridPagos);
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void frm_consulta_planificacion_pagos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void uGridPagos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridPagos);
		}

		private void uGridPagos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Valor", 2);
		}

		private void uGridPagos_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (this.uGridPagos.ActiveCell.Column.Key.ToString().ToUpper() == "VALOR" && (bool)this.uGridPagos.ActiveRow.Cells["VMA"].Value)
				e.Handled=true;
		}

		private void uGridPagos_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if(e.Cell.Column.ToString() == "VMA")
			{
				this.uGridPagos.UpdateData();

				if((bool)e.Cell.Row.Cells["VMA"].Value)
					e.Cell.Row.Cells["Valor"].Value = 0.00m;
			}
		}

		private void uGridPagos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "DiaEmision")	
			{
				if (e.Cell.Row.Cells["DiaEmision"].Value == System.DBNull.Value) e.Cell.Row.Cells["DiaEmision"].Value = (int)e.Cell.OriginalValue;

				if ((int)e.Cell.Row.Cells["DiaEmision"].Value > 30) e.Cell.Row.Cells["DiaEmision"].Value = (int)e.Cell.OriginalValue;
			}
			if (e.Cell.Column.ToString() == "DiaCancelacion")
			{
				if (e.Cell.Row.Cells["DiaCancelacion"].Value == System.DBNull.Value) e.Cell.Row.Cells["DiaCancelacion"].Value = (int)e.Cell.OriginalValue;

				if ((int)e.Cell.Row.Cells["DiaCancelacion"].Value > 30) e.Cell.Row.Cells["DiaCancelacion"].Value = (int)e.Cell.OriginalValue;			
			}
			if (e.Cell.Column.ToString() == "Valor")	
				if (e.Cell.Row.Cells["Valor"].Value == System.DBNull.Value) e.Cell.Row.Cells["Valor"].Value = (decimal)e.Cell.OriginalValue;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void uGridPagos_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			for (int i = 0; i < e.Rows.Length; i++)
			{
				if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de borrar las lineas seleccionadas?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
				{
					if ((int)e.Rows[i].Cells["idPlanificacionPagos"].Value > 0)
					{
						string sSQL = string.Format("Delete From PlanificacionPagos Where idPlanificacionPagos = {0}", (int)e.Rows[i].Cells["idPlanificacionPagos"].Value);
						Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

						e.DisplayPromptMsg = false;
					}
					else e.DisplayPromptMsg = false;
				}
				else e.Cancel = true;
			}
		}		
	}
}
