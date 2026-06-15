using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de ConfitecaIngresaPC.
	/// </summary>
	public class ConfitecaIngresaPC : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtEmpresa;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDireccion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTelefono;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel3;
		private Infragistics.Win.Misc.UltraLabel ultraLabel4;
		private Infragistics.Win.Misc.UltraLabel ultraLabel5;
		private Infragistics.Win.Misc.UltraLabel ultraLabel6;
		private Infragistics.Win.Misc.UltraLabel ultraLabel7;
		private Infragistics.Win.Misc.UltraLabel ultraLabel8;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNombre;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCodigo;
		private C1.Data.C1DataSet cdsConfitecaComision;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private C1.Data.C1DataSet cdsConfiteca;
		private C1.Data.C1DataView cdvConfiteca;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorComisionRef;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorComision;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtIdPadre;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkReferido;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtIdConfitecaComision;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public ConfitecaIngresaPC()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("P_ConfitecaComision", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idComisionConfiteca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Empresa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorccomisionRef");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Referido");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("P_ConfitecaComision", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idComisionConfiteca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Empresa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorccomisionRef");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Referido");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsConfitecaComision = new C1.Data.C1DataSet();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtEmpresa = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDireccion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtTelefono = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel3 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel4 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel5 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel6 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel7 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel8 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbNombre = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvConfiteca = new C1.Data.C1DataView();
			this.cdsConfiteca = new C1.Data.C1DataSet();
			this.cmbCodigo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.txtPorComisionRef = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPorComision = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdPadre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.chkReferido = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.txtIdConfitecaComision = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsConfitecaComision)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvConfiteca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsConfiteca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorComisionRef)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorComision)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdPadre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdConfitecaComision)).BeginInit();
			this.SuspendLayout();
			// 
			// txtNombre
			// 
			this.txtNombre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtNombre.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsConfitecaComision, "P_ConfitecaComision.Nombre"));
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(88, 72);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(232, 21);
			this.txtNombre.TabIndex = 6;
			// 
			// cdsConfitecaComision
			// 
			this.cdsConfitecaComision.BindingContextCtrl = this;
			this.cdsConfitecaComision.CaseSensitive = false;
			this.cdsConfitecaComision.DataLibrary = "LibFacturacion";
			this.cdsConfitecaComision.DataLibraryUrl = "";
			this.cdsConfitecaComision.DataSetDef = "dsConfitecaComision";
			this.cdsConfitecaComision.Locale = new System.Globalization.CultureInfo("es-ES");
			this.cdsConfitecaComision.Name = "cdsConfitecaComision";
			this.cdsConfitecaComision.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsConfitecaComision.SchemaDef = null;
			this.cdsConfitecaComision.CurrentRowChanged += new C1.Data.CurrentRowChangedEventHandler(this.cdsConfitecaComision_CurrentRowChanged);
			this.cdsConfitecaComision.BeforeFill += new C1.Data.FillEventHandler(this.cdsConfitecaComision_BeforeFill);
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(16, 16);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(40, 14);
			this.ultraLabel2.TabIndex = 7;
			this.ultraLabel2.Text = "Codigo";
			// 
			// txtCodigo
			// 
			this.txtCodigo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtCodigo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsConfitecaComision, "P_ConfitecaComision.Codigo"));
			this.txtCodigo.Enabled = false;
			this.txtCodigo.Location = new System.Drawing.Point(88, 16);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(88, 21);
			this.txtCodigo.TabIndex = 8;
			// 
			// txtEmpresa
			// 
			this.txtEmpresa.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtEmpresa.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsConfitecaComision, "P_ConfitecaComision.Empresa"));
			this.txtEmpresa.Enabled = false;
			this.txtEmpresa.Location = new System.Drawing.Point(88, 120);
			this.txtEmpresa.Name = "txtEmpresa";
			this.txtEmpresa.Size = new System.Drawing.Size(232, 21);
			this.txtEmpresa.TabIndex = 9;
			// 
			// txtDireccion
			// 
			this.txtDireccion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtDireccion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsConfitecaComision, "P_ConfitecaComision.Direccion"));
			this.txtDireccion.Enabled = false;
			this.txtDireccion.Location = new System.Drawing.Point(88, 168);
			this.txtDireccion.Name = "txtDireccion";
			this.txtDireccion.Size = new System.Drawing.Size(424, 21);
			this.txtDireccion.TabIndex = 10;
			// 
			// txtTelefono
			// 
			this.txtTelefono.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtTelefono.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsConfitecaComision, "P_ConfitecaComision.Telefono"));
			this.txtTelefono.Enabled = false;
			this.txtTelefono.Location = new System.Drawing.Point(400, 120);
			this.txtTelefono.Name = "txtTelefono";
			this.txtTelefono.Size = new System.Drawing.Size(112, 21);
			this.txtTelefono.TabIndex = 11;
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(344, 120);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(48, 14);
			this.ultraLabel1.TabIndex = 12;
			this.ultraLabel1.Text = "Telefono";
			// 
			// ultraLabel3
			// 
			this.ultraLabel3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.ultraLabel3.AutoSize = true;
			this.ultraLabel3.Location = new System.Drawing.Point(16, 168);
			this.ultraLabel3.Name = "ultraLabel3";
			this.ultraLabel3.Size = new System.Drawing.Size(52, 14);
			this.ultraLabel3.TabIndex = 13;
			this.ultraLabel3.Text = "Direccion";
			// 
			// ultraLabel4
			// 
			this.ultraLabel4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.ultraLabel4.AutoSize = true;
			this.ultraLabel4.Location = new System.Drawing.Point(16, 72);
			this.ultraLabel4.Name = "ultraLabel4";
			this.ultraLabel4.Size = new System.Drawing.Size(44, 14);
			this.ultraLabel4.TabIndex = 14;
			this.ultraLabel4.Text = "Nombre";
			// 
			// ultraLabel5
			// 
			this.ultraLabel5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.ultraLabel5.AutoSize = true;
			this.ultraLabel5.Location = new System.Drawing.Point(16, 120);
			this.ultraLabel5.Name = "ultraLabel5";
			this.ultraLabel5.Size = new System.Drawing.Size(49, 14);
			this.ultraLabel5.TabIndex = 15;
			this.ultraLabel5.Text = "Empresa";
			// 
			// ultraLabel6
			// 
			this.ultraLabel6.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.ultraLabel6.AutoSize = true;
			this.ultraLabel6.Location = new System.Drawing.Point(16, 216);
			this.ultraLabel6.Name = "ultraLabel6";
			this.ultraLabel6.Size = new System.Drawing.Size(65, 14);
			this.ultraLabel6.TabIndex = 18;
			this.ultraLabel6.Text = "% Comision";
			// 
			// ultraLabel7
			// 
			this.ultraLabel7.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.ultraLabel7.AutoSize = true;
			this.ultraLabel7.Location = new System.Drawing.Point(280, 216);
			this.ultraLabel7.Name = "ultraLabel7";
			this.ultraLabel7.Size = new System.Drawing.Size(136, 14);
			this.ultraLabel7.TabIndex = 19;
			this.ultraLabel7.Text = "% Comision por Referidos";
			// 
			// ultraLabel8
			// 
			this.ultraLabel8.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.ultraLabel8.AutoSize = true;
			this.ultraLabel8.Location = new System.Drawing.Point(16, 280);
			this.ultraLabel8.Name = "ultraLabel8";
			this.ultraLabel8.Size = new System.Drawing.Size(71, 14);
			this.ultraLabel8.TabIndex = 21;
			this.ultraLabel8.Text = "Referido Por:";
			// 
			// cmbNombre
			// 
			this.cmbNombre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.cmbNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNombre.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNombre.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsConfitecaComision, "P_ConfitecaComision.idPadre"));
			this.cmbNombre.DataSource = this.cdvConfiteca;
			this.cmbNombre.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "Articulo";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 77;
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Width = 113;
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn3.Width = 101;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 61;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 61;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 61;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 53;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 61;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 34;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 43;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			this.cmbNombre.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbNombre.DisplayMember = "Nombre";
			this.cmbNombre.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNombre.Enabled = false;
			this.cmbNombre.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNombre.Location = new System.Drawing.Point(88, 280);
			this.cmbNombre.Name = "cmbNombre";
			this.cmbNombre.Size = new System.Drawing.Size(232, 21);
			this.cmbNombre.TabIndex = 22;
			this.cmbNombre.ValueMember = "idComisionConfiteca";
			// 
			// cdvConfiteca
			// 
			this.cdvConfiteca.BindingContextCtrl = this;
			this.cdvConfiteca.DataSet = this.cdsConfiteca;
			this.cdvConfiteca.RowFilter = "Referido <= 0";
			this.cdvConfiteca.Sort = "Nombre";
			this.cdvConfiteca.TableName = "";
			this.cdvConfiteca.TableViewName = "P_ConfitecaComision";
			// 
			// cdsConfiteca
			// 
			this.cdsConfiteca.BindingContextCtrl = this;
			this.cdsConfiteca.CaseSensitive = false;
			this.cdsConfiteca.DataLibrary = "LibFacturacion";
			this.cdsConfiteca.DataLibraryUrl = "";
			this.cdsConfiteca.DataSetDef = "dsConfitecaComision";
			this.cdsConfiteca.Locale = new System.Globalization.CultureInfo("es-ES");
			this.cdsConfiteca.Name = "cdsConfiteca";
			this.cdsConfiteca.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsConfiteca.SchemaDef = null;
			this.cdsConfiteca.BeforeFill += new C1.Data.FillEventHandler(this.cdsConfiteca_BeforeFill);
			// 
			// cmbCodigo
			// 
			this.cmbCodigo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.cmbCodigo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCodigo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCodigo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsConfitecaComision, "P_ConfitecaComision.idPadre"));
			this.cmbCodigo.DataSource = this.cdvConfiteca;
			this.cmbCodigo.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "Articulo";
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 79;
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Width = 87;
			ultraGridColumn13.Header.VisiblePosition = 2;
			ultraGridColumn13.Width = 79;
			ultraGridColumn14.Header.VisiblePosition = 3;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn14.Width = 58;
			ultraGridColumn15.Header.VisiblePosition = 4;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Width = 58;
			ultraGridColumn16.Header.VisiblePosition = 5;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn16.Width = 58;
			ultraGridColumn17.Header.VisiblePosition = 6;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn17.Width = 51;
			ultraGridColumn18.Header.VisiblePosition = 7;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn18.Width = 62;
			ultraGridColumn19.Header.VisiblePosition = 8;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Width = 33;
			ultraGridColumn20.Header.VisiblePosition = 9;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn20.Width = 58;
			ultraGridBand2.Columns.AddRange(new object[] {
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
			this.cmbCodigo.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCodigo.DisplayMember = "Codigo";
			this.cmbCodigo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCodigo.Enabled = false;
			this.cmbCodigo.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCodigo.Location = new System.Drawing.Point(328, 280);
			this.cmbCodigo.Name = "cmbCodigo";
			this.cmbCodigo.Size = new System.Drawing.Size(184, 21);
			this.cmbCodigo.TabIndex = 23;
			this.cmbCodigo.ValueMember = "idComisionConfiteca";
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "P_ConfitecaComision";
			this.barraDato1.DataNombreId = "idComisionConfiteca";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsConfitecaComision;
			this.barraDato1.DataTabla = "P_ConfitecaComision";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Location = new System.Drawing.Point(16, 320);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 24);
			this.barraDato1.TabIndex = 24;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = false;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = false;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Restaurar += new System.EventHandler(this.barraDato1_Restaurar);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// txtPorComisionRef
			// 
			this.txtPorComisionRef.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPorComisionRef.Appearance = appearance1;
			this.txtPorComisionRef.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsConfitecaComision, "P_ConfitecaComision.PorccomisionRef"));
			this.txtPorComisionRef.Enabled = false;
			this.txtPorComisionRef.FormatString = "#,##0.00";
			this.txtPorComisionRef.Location = new System.Drawing.Point(432, 216);
			this.txtPorComisionRef.Name = "txtPorComisionRef";
			this.txtPorComisionRef.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorComisionRef.PromptChar = ' ';
			this.txtPorComisionRef.Size = new System.Drawing.Size(80, 21);
			this.txtPorComisionRef.TabIndex = 130;
			// 
			// txtPorComision
			// 
			this.txtPorComision.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPorComision.Appearance = appearance2;
			this.txtPorComision.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsConfitecaComision, "P_ConfitecaComision.PorcComision"));
			this.txtPorComision.Enabled = false;
			this.txtPorComision.FormatString = "#,##0.00";
			this.txtPorComision.Location = new System.Drawing.Point(88, 216);
			this.txtPorComision.Name = "txtPorComision";
			this.txtPorComision.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorComision.PromptChar = ' ';
			this.txtPorComision.Size = new System.Drawing.Size(80, 21);
			this.txtPorComision.TabIndex = 131;
			// 
			// txtIdPadre
			// 
			this.txtIdPadre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtIdPadre.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsConfitecaComision, "P_ConfitecaComision.idPadre"));
			this.txtIdPadre.Enabled = false;
			this.txtIdPadre.Location = new System.Drawing.Point(88, 16);
			this.txtIdPadre.Name = "txtIdPadre";
			this.txtIdPadre.Size = new System.Drawing.Size(32, 21);
			this.txtIdPadre.TabIndex = 133;
			// 
			// chkReferido
			// 
			this.chkReferido.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkReferido.Appearance = appearance3;
			this.chkReferido.BackColor = System.Drawing.SystemColors.Control;
			this.chkReferido.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsConfitecaComision, "P_ConfitecaComision.Referido"));
			this.chkReferido.Enabled = false;
			this.chkReferido.Location = new System.Drawing.Point(216, 16);
			this.chkReferido.Name = "chkReferido";
			this.chkReferido.Size = new System.Drawing.Size(88, 20);
			this.chkReferido.TabIndex = 132;
			this.chkReferido.Text = "Referido";
			this.chkReferido.CheckedChanged += new System.EventHandler(this.chkReferido_CheckedChanged);
			// 
			// txtIdConfitecaComision
			// 
			this.txtIdConfitecaComision.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtIdConfitecaComision.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsConfitecaComision, "P_ConfitecaComision.idComisionConfiteca"));
			this.txtIdConfitecaComision.Enabled = false;
			this.txtIdConfitecaComision.Location = new System.Drawing.Point(128, 16);
			this.txtIdConfitecaComision.Name = "txtIdConfitecaComision";
			this.txtIdConfitecaComision.Size = new System.Drawing.Size(32, 21);
			this.txtIdConfitecaComision.TabIndex = 134;
			// 
			// ConfitecaIngresaPC
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(552, 365);
			this.Controls.Add(this.chkReferido);
			this.Controls.Add(this.txtPorComision);
			this.Controls.Add(this.txtPorComisionRef);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.cmbCodigo);
			this.Controls.Add(this.cmbNombre);
			this.Controls.Add(this.ultraLabel8);
			this.Controls.Add(this.ultraLabel7);
			this.Controls.Add(this.ultraLabel6);
			this.Controls.Add(this.ultraLabel5);
			this.Controls.Add(this.ultraLabel4);
			this.Controls.Add(this.ultraLabel3);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.txtTelefono);
			this.Controls.Add(this.txtDireccion);
			this.Controls.Add(this.txtEmpresa);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.txtIdPadre);
			this.Controls.Add(this.txtIdConfitecaComision);
			this.Name = "ConfitecaIngresaPC";
			this.Text = "Ingreso de Personal a Comisionar";
			this.Load += new System.EventHandler(this.ConfitecaIngresaPC_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsConfitecaComision)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvConfiteca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsConfiteca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorComisionRef)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorComision)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdPadre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdConfitecaComision)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
//C1.Data.C1DataRow drConfiteca;
		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
		barraDato1.EditarRegistro();
			controla();
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
		barraDato1.GrabaRegistro();
			this.cdsConfiteca.Clear();
			this.cdsConfiteca.Fill();		
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{

			if (DialogResult.No == MessageBox.Show("¿Desea borrar registro?", "Confirmación", 
				MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				return;

			string stSelect = string.Format( "select COUNT(*) from compra c inner join p_ConfitecaComision cc on c.idComisionConfiteca = cc.idComisionConfiteca "+
											"where c.idComisionConfiteca = {0} or {0} = cc.idPadre",txtIdConfitecaComision.Value.ToString());

			int cuenta = Funcion.iEscalarSQL(cdsConfitecaComision,stSelect);
			if(cuenta>0)
			{
				MessageBox.Show("Existen "+cuenta+" registros relacionados","No Puede borrar ",MessageBoxButtons.OK,MessageBoxIcon.Information);
				return;
			}
		barraDato1.BorraRegistro();
			this.cdsConfiteca.Clear();
			this.cdsConfiteca.Fill();	
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			barraDato1.CrearRegistro();
			//drConfiteca = cdsConfitecaComision.TableViews["P_ConfitecaComision"].Rows[0]; 

//			chkReferido.Checked=false;
//			txtPorComision.Value=0;
//			txtPorComisionRef.Value=0;
			
//			this.drConfiteca["Codigo"]=0;
//			this.drConfiteca["Nombre"]=0;
//			this.drConfiteca["Empresa"]=0;
//			this.drConfiteca["Direccion"]=0;
//			this.drConfiteca["Telefono"]=0;
//			this.drConfiteca["Porccomision"]=0;
//			this.drConfiteca["PorccomisionRef"]=0;
//			this.drConfiteca["idPadre"]=0;
//			this.drConfiteca["Referido"]=true;
//			cmbNombre.Value = 0;
//			cmbCodigo.Value=0;
						
		}

		private void ConfitecaIngresaPC_Load(object sender, System.EventArgs e)
		{			
			barraDato1.PosUltima();				
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{				
			barraDato1.DeshacerRegistro();
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
       
		}

		private void cdsConfitecaComision_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
				Funcion miFuncion = new Funcion();
			this.cdsConfitecaComision.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);		
		}

		private void cdsConfiteca_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
//			Funcion miFuncion = new Funcion();
//			this.cdsConfiteca.Schema.Connections[0].ConnectionString = 
//				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}
		private void chkReferido_CheckedChanged(object sender, System.EventArgs e)
		{
//			return;
			controla();
		}

		private void controla()
		{
			//drConfiteca = cdsConfitecaComision.TableViews["P_ConfitecaComision"].Rows[0];
			if (chkReferido.Enabled == true)
			{
				//if (drConfiteca == null) return;

				if(chkReferido.Checked == true)
				{
					txtPorComisionRef.Enabled=false;
					cmbCodigo.Enabled = true;
					cmbNombre.Enabled = true;				
					//drConfiteca["PorccomisionRef"]=0;
					txtPorComisionRef.Value = 0;					

				}
				if(chkReferido.Checked == false)
				{					
					txtPorComisionRef.Enabled=true;
					cmbCodigo.Enabled = false;
					cmbNombre.Enabled = false;
					//drConfiteca["idPadre"]=0;         
					txtIdPadre.Value=0;
				}
			}
		}
		private void cdsConfitecaComision_CurrentRowChanged(object sender, C1.Data.CurrentRowChangedEventArgs e)
		{
//		drConfiteca = cdsConfitecaComision.TableViews["P_ConfitecaComision"].Rows[0]; 		
		}

		private void barraDato1_Restaurar(object sender, System.EventArgs e)
		{
		//drConfiteca = cdsConfitecaComision.TableViews["P_ConfitecaComision"].Rows[0]; 		
		}


		
	
	}
}
