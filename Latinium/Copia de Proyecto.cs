using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Proyecto.
	/// </summary>
	public class Proyecto : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsProyecto;
		private Infragistics.Win.Misc.UltraButton btSubProyecto;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private Infragistics.Win.Misc.UltraButton btAuditoria;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCentro;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor2;
		private Infragistics.Win.Misc.UltraLabel ultraLabel4;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor4;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel5;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtIdProyecto;
		private C1.Data.C1DataSet cdsSeguridad;
		private DevExpress.XtraEditors.LookUpEdit cmbDistribucion;
		private Infragistics.Win.Misc.UltraLabel ultraLabel6;
		private DevExpress.XtraEditors.SimpleButton btDist;
		private System.Windows.Forms.Label lblSubTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCajaChica;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCuentas;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Proyecto()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: Agregar código de constructor después de llamar a InitializeComponent
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Método necesario para admitir el Diseñador, no se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(Proyecto));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			this.cdsProyecto = new C1.Data.C1DataSet();
			this.btSubProyecto = new Infragistics.Win.Misc.UltraButton();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.btAuditoria = new Infragistics.Win.Misc.UltraButton();
			this.txtCentro = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraTextEditor1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel3 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraTextEditor2 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel4 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraTextEditor3 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraTextEditor4 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraCheckEditor1 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraLabel5 = new Infragistics.Win.Misc.UltraLabel();
			this.txtIdProyecto = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			this.cmbDistribucion = new DevExpress.XtraEditors.LookUpEdit();
			this.ultraLabel6 = new Infragistics.Win.Misc.UltraLabel();
			this.btDist = new DevExpress.XtraEditors.SimpleButton();
			this.lblSubTotal = new System.Windows.Forms.Label();
			this.txtCajaChica = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbCuentas = new Infragistics.Win.UltraWinGrid.UltraCombo();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCentro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDistribucion.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCajaChica)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentas)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsProyecto
			// 
			this.cdsProyecto.BindingContextCtrl = this;
			this.cdsProyecto.DataLibrary = "LibContabilidad";
			this.cdsProyecto.DataLibraryUrl = "";
			this.cdsProyecto.DataSetDef = "dsProyecto";
			this.cdsProyecto.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsProyecto.Name = "cdsProyecto";
			this.cdsProyecto.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsProyecto.SchemaDef = null;
			this.cdsProyecto.BeforeFill += new C1.Data.FillEventHandler(this.cdsProyecto_BeforeFill);
			// 
			// btSubProyecto
			// 
			this.btSubProyecto.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btSubProyecto.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btSubProyecto.Location = new System.Drawing.Point(392, 240);
			this.btSubProyecto.Name = "btSubProyecto";
			this.btSubProyecto.Size = new System.Drawing.Size(80, 25);
			this.btSubProyecto.TabIndex = 3;
			this.btSubProyecto.Text = "&Proyecto";
			this.btSubProyecto.Click += new System.EventHandler(this.btSubProyecto_Click);
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "Proyecto";
			this.barraDato1.DataNombreId = "idProyecto";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsProyecto;
			this.barraDato1.DataTabla = "Proyecto";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Location = new System.Drawing.Point(24, 240);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 26);
			this.barraDato1.TabIndex = 11;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = true;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Buscar += new System.EventHandler(this.barraDato1_Buscar);
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Imprimir += new System.EventHandler(this.barraDato1_Imprimir);
			this.barraDato1.Restaurar += new System.EventHandler(this.barraDato1_Restaurar);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Copiar += new System.EventHandler(this.barraDato1_Copiar);
			this.barraDato1.Impresora += new System.EventHandler(this.barraDato1_Impresora);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// btAuditoria
			// 
			this.btAuditoria.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance1.Image = ((object)(resources.GetObject("appearance1.Image")));
			this.btAuditoria.Appearance = appearance1;
			this.btAuditoria.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAuditoria.Location = new System.Drawing.Point(352, 240);
			this.btAuditoria.Name = "btAuditoria";
			this.btAuditoria.Size = new System.Drawing.Size(24, 26);
			this.btAuditoria.TabIndex = 122;
			this.btAuditoria.Click += new System.EventHandler(this.btAuditoria_Click);
			this.btAuditoria.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btAuditoria_MouseUp);
			// 
			// txtCentro
			// 
			this.txtCentro.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsProyecto, "Proyecto.CentroCosto"));
			this.txtCentro.Location = new System.Drawing.Point(144, 62);
			this.txtCentro.Name = "txtCentro";
			this.txtCentro.Size = new System.Drawing.Size(208, 22);
			this.txtCentro.TabIndex = 123;
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(24, 60);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(38, 15);
			this.ultraLabel1.TabIndex = 124;
			this.ultraLabel1.Text = "Código";
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(24, 26);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(44, 15);
			this.ultraLabel2.TabIndex = 126;
			this.ultraLabel2.Text = "Nombre";
			// 
			// ultraTextEditor1
			// 
			this.ultraTextEditor1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsProyecto, "Proyecto.Nombre"));
			this.ultraTextEditor1.Location = new System.Drawing.Point(144, 26);
			this.ultraTextEditor1.Name = "ultraTextEditor1";
			this.ultraTextEditor1.Size = new System.Drawing.Size(440, 22);
			this.ultraTextEditor1.TabIndex = 125;
			// 
			// ultraLabel3
			// 
			this.ultraLabel3.AutoSize = true;
			this.ultraLabel3.Location = new System.Drawing.Point(24, 129);
			this.ultraLabel3.Name = "ultraLabel3";
			this.ultraLabel3.Size = new System.Drawing.Size(67, 15);
			this.ultraLabel3.TabIndex = 130;
			this.ultraLabel3.Text = "Responsable";
			// 
			// ultraTextEditor2
			// 
			this.ultraTextEditor2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsProyecto, "Proyecto.Responsable"));
			this.ultraTextEditor2.Location = new System.Drawing.Point(144, 134);
			this.ultraTextEditor2.Name = "ultraTextEditor2";
			this.ultraTextEditor2.Size = new System.Drawing.Size(208, 22);
			this.ultraTextEditor2.TabIndex = 129;
			// 
			// ultraLabel4
			// 
			this.ultraLabel4.AutoSize = true;
			this.ultraLabel4.Location = new System.Drawing.Point(24, 95);
			this.ultraLabel4.Name = "ultraLabel4";
			this.ultraLabel4.Size = new System.Drawing.Size(32, 15);
			this.ultraLabel4.TabIndex = 128;
			this.ultraLabel4.Text = "Lugar";
			// 
			// ultraTextEditor3
			// 
			this.ultraTextEditor3.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsProyecto, "Proyecto.Lugar"));
			this.ultraTextEditor3.Location = new System.Drawing.Point(144, 98);
			this.ultraTextEditor3.Name = "ultraTextEditor3";
			this.ultraTextEditor3.Size = new System.Drawing.Size(208, 22);
			this.ultraTextEditor3.TabIndex = 127;
			// 
			// ultraTextEditor4
			// 
			this.ultraTextEditor4.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsProyecto, "Proyecto.Notas"));
			this.ultraTextEditor4.Location = new System.Drawing.Point(376, 80);
			this.ultraTextEditor4.Multiline = true;
			this.ultraTextEditor4.Name = "ultraTextEditor4";
			this.ultraTextEditor4.Size = new System.Drawing.Size(208, 104);
			this.ultraTextEditor4.TabIndex = 131;
			// 
			// ultraCheckEditor1
			// 
			this.ultraCheckEditor1.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsProyecto, "Proyecto.Visible"));
			this.ultraCheckEditor1.Location = new System.Drawing.Point(496, 56);
			this.ultraCheckEditor1.Name = "ultraCheckEditor1";
			this.ultraCheckEditor1.Size = new System.Drawing.Size(72, 17);
			this.ultraCheckEditor1.TabIndex = 132;
			this.ultraCheckEditor1.Text = "Visible";
			// 
			// ultraLabel5
			// 
			this.ultraLabel5.Location = new System.Drawing.Point(376, 56);
			this.ultraLabel5.Name = "ultraLabel5";
			this.ultraLabel5.Size = new System.Drawing.Size(48, 17);
			this.ultraLabel5.TabIndex = 133;
			this.ultraLabel5.Text = "Notas";
			// 
			// txtIdProyecto
			// 
			this.txtIdProyecto.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsProyecto, "Proyecto.idProyecto"));
			this.txtIdProyecto.Location = new System.Drawing.Point(536, 240);
			this.txtIdProyecto.Name = "txtIdProyecto";
			this.txtIdProyecto.Size = new System.Drawing.Size(56, 22);
			this.txtIdProyecto.TabIndex = 134;
			// 
			// cdsSeguridad
			// 
			this.cdsSeguridad.BindingContextCtrl = this;
			this.cdsSeguridad.DataLibrary = "LibContabilidad";
			this.cdsSeguridad.DataLibraryUrl = "";
			this.cdsSeguridad.DataSetDef = "dsSeguridad";
			this.cdsSeguridad.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeguridad.Name = "cdsSeguridad";
			this.cdsSeguridad.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeguridad.SchemaDef = null;
			// 
			// cmbDistribucion
			// 
			this.cmbDistribucion.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsProyecto, "Proyecto.idAsiDistGrupo"));
			this.cmbDistribucion.Location = new System.Drawing.Point(144, 170);
			this.cmbDistribucion.Name = "cmbDistribucion";
			// 
			// cmbDistribucion.Properties
			// 
			this.cmbDistribucion.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																						new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbDistribucion.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																								new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Distribucion", "Distribucion", 49, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near, DevExpress.Data.ColumnSortOrder.Ascending)});
			this.cmbDistribucion.Properties.DisplayMember = "Distribucion";
			this.cmbDistribucion.Properties.NullText = "";
			this.cmbDistribucion.Properties.PopupWidth = 300;
			this.cmbDistribucion.Properties.ShowFooter = false;
			this.cmbDistribucion.Properties.ShowHeader = false;
			this.cmbDistribucion.Properties.ValueMember = "idAsiDistGrupo";
			this.cmbDistribucion.Size = new System.Drawing.Size(176, 20);
			this.cmbDistribucion.TabIndex = 135;
			this.cmbDistribucion.MouseDown += new System.Windows.Forms.MouseEventHandler(this.cmbDistribucion_MouseDown);
			// 
			// ultraLabel6
			// 
			this.ultraLabel6.AutoSize = true;
			this.ultraLabel6.Location = new System.Drawing.Point(24, 168);
			this.ultraLabel6.Name = "ultraLabel6";
			this.ultraLabel6.Size = new System.Drawing.Size(63, 15);
			this.ultraLabel6.TabIndex = 136;
			this.ultraLabel6.Text = "Distribución";
			// 
			// btDist
			// 
			this.btDist.Location = new System.Drawing.Point(328, 168);
			this.btDist.Name = "btDist";
			this.btDist.Size = new System.Drawing.Size(24, 23);
			this.btDist.TabIndex = 137;
			this.btDist.Text = "...";
			this.btDist.Click += new System.EventHandler(this.btDist_Click);
			this.btDist.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btDist_MouseUp);
			// 
			// lblSubTotal
			// 
			this.lblSubTotal.AutoSize = true;
			this.lblSubTotal.Location = new System.Drawing.Point(24, 208);
			this.lblSubTotal.Name = "lblSubTotal";
			this.lblSubTotal.Size = new System.Drawing.Size(56, 16);
			this.lblSubTotal.TabIndex = 138;
			this.lblSubTotal.Text = "Caja Chica";
			// 
			// txtCajaChica
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCajaChica.Appearance = appearance2;
			this.txtCajaChica.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsProyecto, "Proyecto.CajaChica"));
			this.txtCajaChica.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCajaChica.FormatString = "#,##0.00";
			this.txtCajaChica.Location = new System.Drawing.Point(144, 204);
			this.txtCajaChica.Name = "txtCajaChica";
			this.txtCajaChica.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCajaChica.PromptChar = ' ';
			this.txtCajaChica.Size = new System.Drawing.Size(80, 22);
			this.txtCajaChica.TabIndex = 139;
			this.txtCajaChica.Click += new System.EventHandler(this.txtCajaChica_Click);
			this.txtCajaChica.Enter += new System.EventHandler(this.txtCajaChica_Enter);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(264, 208);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(39, 16);
			this.label4.TabIndex = 305;
			this.label4.Text = "Cuenta";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// cmbCuentas
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCuentas.Appearance = appearance3;
			this.cmbCuentas.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuentas.Cursor = System.Windows.Forms.Cursors.WaitCursor;
			this.cmbCuentas.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsProyecto, "Proyecto.idCuenta"));
			this.cmbCuentas.DataSource = this.ultraDataSource1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 208;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbCuentas.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCuentas.DisplayMember = "Descripcion";
			this.cmbCuentas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuentas.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCuentas.Enabled = false;
			this.cmbCuentas.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentas.Location = new System.Drawing.Point(328, 208);
			this.cmbCuentas.Name = "cmbCuentas";
			this.cmbCuentas.Size = new System.Drawing.Size(256, 21);
			this.cmbCuentas.TabIndex = 306;
			this.cmbCuentas.ValueMember = "idCuenta";
			// 
			// Proyecto
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(600, 270);
			this.Controls.Add(this.cmbCuentas);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.lblSubTotal);
			this.Controls.Add(this.txtCajaChica);
			this.Controls.Add(this.btDist);
			this.Controls.Add(this.ultraLabel6);
			this.Controls.Add(this.cmbDistribucion);
			this.Controls.Add(this.txtIdProyecto);
			this.Controls.Add(this.ultraLabel5);
			this.Controls.Add(this.ultraCheckEditor1);
			this.Controls.Add(this.ultraTextEditor4);
			this.Controls.Add(this.ultraLabel3);
			this.Controls.Add(this.ultraTextEditor2);
			this.Controls.Add(this.ultraLabel4);
			this.Controls.Add(this.ultraTextEditor3);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.ultraTextEditor1);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.txtCentro);
			this.Controls.Add(this.btAuditoria);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.btSubProyecto);
			this.Name = "Proyecto";
			this.Text = "Centros de Costo";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.Proyecto_Closing);
			this.Load += new System.EventHandler(this.Proyecto_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCentro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDistribucion.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCajaChica)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentas)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			try
			{
				this.cdsProyecto.Update();
				this.Close();
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Información");
			}
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
//			switch(e.KeyValue)
//			{
//				case (int) Keys.PageDown:
//					if(grProyecto.ActiveCell.DroppedDown == false)
//						grProyecto.ActiveCell.DroppedDown = true;
//					break;
//
//				case (int) Keys.Up:
//					if (grProyecto.ActiveCell.DroppedDown) return;
//
//					grProyecto.PerformAction(UltraGridAction.ExitEditMode, false, false);
//					grProyecto.PerformAction(UltraGridAction.AboveCell, false, false);
//					e.Handled = true;
//					grProyecto.PerformAction(UltraGridAction.EnterEditMode, false, false);
//					break;
//				case (int) Keys.Down:
//					if (grProyecto.ActiveCell.DroppedDown) return;
//			
//					grProyecto.PerformAction(UltraGridAction.ExitEditMode, false, false);
//					grProyecto.PerformAction(UltraGridAction.BelowCell, false, false);
//					e.Handled = true;
//					grProyecto.PerformAction(UltraGridAction.EnterEditMode, false, false);
//					break;
//				case (int) Keys.Enter:
//					grProyecto.PerformAction(UltraGridAction.ExitEditMode, false, false);
//					grProyecto.PerformAction(UltraGridAction.NextCellByTab, false, false);
//					e.Handled = true;
//					grProyecto.PerformAction(UltraGridAction.EnterEditMode, false, false);
//					break;
//			}
//
		}

		private void btSubProyecto_Click(object sender, System.EventArgs e)
		{
			SubProyecto miSubProyecto = new SubProyecto();
			miSubProyecto.MdiParent = this.MdiParent;
			miSubProyecto.Show();
		}

		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}
		private void Proyecto_Load(object sender, System.EventArgs e)
		{
			if (!Funcion.Permiso("63",cdsProyecto))
			{
				MessageBox.Show("No puede ingresar a Centro de Costo", "Informacion");
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			string stAudita = "Exec AuditaCrear 53, 6, 'Proyecto'";
			Funcion.EjecutaSQL(cdsProyecto, stAudita, true);

			this.cmbCuentas.DataSource = Funcion.dvProcedimiento(cdsProyecto, "Select idCuenta, Codigo, Descripcion From Cuenta Where Padre = 449");

			cmbDistribucion.Properties.DataSource = Funcion.dvProcedimiento(cdsProyecto, "Select idAsiDistGrupo, Distribucion From AsientoDistGrupo");
			miAcceso =new Acceso("0421", cdsSeguridad.StorageDataSet.Tables["Seguridad"]);
			txtIdProyecto.Width = 0;
			Funcion.EjecutaSQL(cdsProyecto, stAudita, true);
			barraDato1.ProximoId(4);
			barraDato1.HabilitaControles(false);
			Cursor = Cursors.Default;
		}

		private void grProyecto_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (DialogResult.No ==	MessageBox.Show("¿Desea Borrar Registro?",
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				e.Cancel = true;
			else 
			{
				for (int i=0; i<e.Rows.Length; i++)
				{
					string stProy = e.Rows[i].Cells["idProyecto"].Value.ToString();
					string stCentro = e.Rows[i].Cells["Nombre"].Value.ToString();
					string stSelect = "Select COUNT(*) FROM Asiento WHERE idProyecto = " + stProy;
					int iCuentaAsi = Funcion.iEscalarSQL(cdsProyecto, stSelect);
					string stMensaje = "No puede Borrar Centro de Costo: " 
						+ stCentro + " porque tiene ";
					if (iCuentaAsi > 0) 
						stMensaje += iCuentaAsi.ToString() + " Asientos. ";

					stSelect = "Select COUNT(*) FROM Compra WHERE idProyecto = " + stProy;
					int iCuentaFac = Funcion.iEscalarSQL(cdsProyecto, stSelect);
					if (iCuentaFac > 0) stMensaje += iCuentaFac.ToString() + " Facturas. ";
					stMensaje += " Relacionados";
					if (iCuentaAsi + iCuentaFac > 0) 
					{
						MessageBox.Show(stMensaje, "Información", MessageBoxButtons.OK,
							MessageBoxIcon.Information);
						e.Cancel = true;
						return;
					}
				}
				e.DisplayPromptMsg = false;
			}
		}

		private void cdsProyecto_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsProyecto.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void Proyecto_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Proyecto'";
			Funcion.EjecutaSQL(cdsProyecto, stAudita, true);
		}

		private void btAuditoria_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			Funcion.AuditoriaImprime(60, int.Parse(txtIdProyecto.Value.ToString()), true);
			Cursor = Cursors.Default;
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Eliminar)
			{
				MessageBox.Show("No tiene acceso a Eliminar Articulos",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (MessageBox.Show("¿Desea Borrar este Proyecto?", "Confirmación",
				MessageBoxButtons.YesNo, MessageBoxIcon.Warning,
				MessageBoxDefaultButton.Button2) == DialogResult.No) 
				return;

			if (this.txtIdProyecto.Value == System.DBNull.Value) return;

			// Verificacion de Existencia de articulos en compras
			try
			{
				string stSelect = "SELECT COUNT(*) FROM Compra WHERE Compra.idProyecto = "
					+ this.txtIdProyecto.Value.ToString().Trim();
				int iCuenta = Funcion.iEscalarSQL(cdsProyecto, stSelect);
				if (iCuenta > 0) 
				{
					MessageBox.Show("No puede borrar este proyecto.\n"
						+"Tiene " + iCuenta.ToString().Trim() + " Transacciones relacionados."
						+ " Eliminación CANCELADA",
						"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
				stSelect = "SELECT COUNT(*) FROM DetCompra WHERE idProyecto = "
					+ this.txtIdProyecto.Value.ToString().Trim();
				iCuenta = Funcion.iEscalarSQL(cdsProyecto, stSelect);
				if (iCuenta > 0) 
				{
					MessageBox.Show("No puede borrar este proyecto.\n"
						+"Tiene " + iCuenta.ToString().Trim() + " Detalles de Transacciones relacionados."
						+ " Eliminación CANCELADA",
						"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
				stSelect = "SELECT COUNT(*) FROM Asiento WHERE Asiento.idProyecto = "
					+ this.txtIdProyecto.Value.ToString().Trim();
				iCuenta = Funcion.iEscalarSQL(cdsProyecto, stSelect);
				if (iCuenta > 0) 
				{
					MessageBox.Show("No puede borrar este proyecto.\n"
						+"Tiene " + iCuenta.ToString().Trim() + " Asientos relacionados."
						+ " Eliminación CANCELADA",
						"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
				stSelect = "SELECT COUNT(*) FROM AsientoDetalle WHERE idProyecto = "
					+ this.txtIdProyecto.Value.ToString().Trim();
				iCuenta = Funcion.iEscalarSQL(cdsProyecto, stSelect);
				if (iCuenta > 0) 
				{
					MessageBox.Show("No puede borrar este proyecto.\n"
						+"Tiene " + iCuenta.ToString().Trim() + " Detalle de Asientos relacionados."
						+ " Eliminación CANCELADA",
						"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Borrar Articulo");
			}
			// Fin de verificacion de Existencia de articulos en compras

			this.barraDato1.BorraRegistro();		
		}

		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
//			if (!miAcceso.Buscar)
//			{
//				MessageBox.Show("No tiene acceso a Buscar Centros de Costo" ,
//					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
//				return;
//			}
			using (ProyectoBusca miBusqueda = new ProyectoBusca())
			{
				if (DialogResult.OK == miBusqueda.ShowDialog())
				{
					try
					{
						int id = int.Parse(miBusqueda.idBusca.Text);
						if (id > 0)
						{
							this.barraDato1.IdRegistro = id;
							this.barraDato1.ProximoId(5);
						}
					}
					catch{}
				}
			}
		
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			if (!miAcceso.Nuevo)
			{
				MessageBox.Show("No tiene acceso a Crear Nuevos Articulos",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			barraDato1.CrearRegistro();
		}

		private void barraDato1_Copiar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			barraDato1.DeshacerRegistro();
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Editar)
			{
				MessageBox.Show("No tiene acceso a Editar Articulos",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			barraDato1.EditarRegistro();
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			barraDato1.GrabaRegistro();
		}

		private void barraDato1_Impresora(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Imprimir(object sender, System.EventArgs e)
		{
			Reporte miRep = new Reporte("Proyectos.Rpt");
			miRep.MdiParent = this.MdiParent;
			miRep.Show();
		}

		private void barraDato1_Restaurar(object sender, System.EventArgs e)
		{
		
		}

		private void btAuditoria_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			Cursor = Cursors.WaitCursor;
			Funcion.AuditoriaImprime(60, int.Parse(txtIdProyecto.Value.ToString()), false);
			Cursor = Cursors.Default;		
		}

		private void btDist_Click(object sender, System.EventArgs e)
		{
			using (AsientoDist miDist = new AsientoDist())
			{
				miDist.ShowDialog();
			}
			cmbDistribucion.Properties.DataSource = Funcion.dvProcedimiento(cdsProyecto,
				"Select idAsiDistGrupo, Distribucion From AsientoDistGrupo");		
		}

		private void cmbDistribucion_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			cmbDistribucion.EditValue = 0;
		}

		private void btDist_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			if (DialogResult.No == MessageBox.Show("Desea Distribuir las cuentas de gasto de todos los Asientos de este proyecto?", 
					"Confirmacion", MessageBoxButtons.YesNo, MessageBoxIcon.Question)	) return;
			string stExec = "Exec ProyectoDistribuyeCC " + txtIdProyecto.Text;
			Funcion.EjecutaSQL(cdsProyecto, stExec);
			MessageBox.Show("Cuentas distribuidas", 
				"Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void txtCajaChica_Enter(object sender, System.EventArgs e)
		{
			this.txtCajaChica.SelectAll();
		}

		private void txtCajaChica_Click(object sender, System.EventArgs e)
		{
			this.txtCajaChica.SelectAll();
		}
	}
}
