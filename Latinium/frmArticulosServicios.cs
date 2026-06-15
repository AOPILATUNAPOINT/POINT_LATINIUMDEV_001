using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmArticulosServicios.
	/// </summary>
	public class frmArticulosServicios : DevExpress.XtraEditors.XtraForm
	{
		private DevExpress.XtraEditors.SimpleButton btArticuloSRI;
		private System.Windows.Forms.Label label46;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbGrupo;
		private DevExpress.XtraEditors.SimpleButton btGrupo;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private C1.Data.C1DataSet cdsArticulo;
		private C1.Data.C1DataSet cdsArticuloTabla;
		private C1.Data.C1DataView cvdArticuloGrupo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtIdArticulo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIva;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label4;
		private C1.Data.C1DataSet cdsSeteoF;
		private Acceso miAcceso;
		private DevExpress.XtraEditors.SimpleButton btLimpia;
		private System.Windows.Forms.Label label5;
		private DevExpress.XtraEditors.SimpleButton btnLimpia1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbRetencionIVA;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbArticuloSRI;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkImpuestoIVA;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmArticulosServicios()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		int idArticulo = 0;
		public frmArticulosServicios(int IdArticulo)
		{
			InitializeComponent();
			idArticulo = IdArticulo;
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmArticulosServicios));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloGrupo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloSRI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas7");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas8");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas9");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr6");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRetIva");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Detalle");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloSRI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticuloSRI");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Detalle");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			this.cdsArticulo = new C1.Data.C1DataSet();
			this.cdsArticuloTabla = new C1.Data.C1DataSet();
			this.btArticuloSRI = new DevExpress.XtraEditors.SimpleButton();
			this.label46 = new System.Windows.Forms.Label();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cvdArticuloGrupo = new C1.Data.C1DataView();
			this.btGrupo = new DevExpress.XtraEditors.SimpleButton();
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.txtIdArticulo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtIva = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label9 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btLimpia = new DevExpress.XtraEditors.SimpleButton();
			this.btnLimpia1 = new DevExpress.XtraEditors.SimpleButton();
			this.label5 = new System.Windows.Forms.Label();
			this.cmbRetencionIVA = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbArticuloSRI = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.chkImpuestoIVA = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticuloTabla)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cvdArticuloGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRetencionIVA)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticuloSRI)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsArticulo
			// 
			this.cdsArticulo.BindingContextCtrl = this;
			this.cdsArticulo.DataLibrary = "LibFacturacion";
			this.cdsArticulo.DataLibraryUrl = "";
			this.cdsArticulo.DataSetDef = "dsArticulo";
			this.cdsArticulo.FillOnRequest = false;
			this.cdsArticulo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsArticulo.Name = "cdsArticulo";
			this.cdsArticulo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsArticulo.SchemaDef = null;
			this.cdsArticulo.BeforeFill += new C1.Data.FillEventHandler(this.cdsArticulo_BeforeFill);
			// 
			// cdsArticuloTabla
			// 
			this.cdsArticuloTabla.BindingContextCtrl = this;
			this.cdsArticuloTabla.CaseSensitive = false;
			this.cdsArticuloTabla.DataLibrary = "LibFacturacion";
			this.cdsArticuloTabla.DataLibraryUrl = "";
			this.cdsArticuloTabla.DataSetDef = "dsArticuloTabla";
			this.cdsArticuloTabla.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsArticuloTabla.Name = "cdsArticuloTabla";
			this.cdsArticuloTabla.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsArticuloTabla.SchemaDef = null;
			this.cdsArticuloTabla.BeforeFill += new C1.Data.FillEventHandler(this.cdsArticuloTabla_BeforeFill);
			// 
			// btArticuloSRI
			// 
			this.btArticuloSRI.Image = ((System.Drawing.Image)(resources.GetObject("btArticuloSRI.Image")));
			this.btArticuloSRI.Location = new System.Drawing.Point(400, 133);
			this.btArticuloSRI.Name = "btArticuloSRI";
			this.btArticuloSRI.Size = new System.Drawing.Size(24, 25);
			this.btArticuloSRI.TabIndex = 115;
			this.btArticuloSRI.ToolTip = "Retenciones de artículos";
			this.btArticuloSRI.Click += new System.EventHandler(this.btArticuloSRI_Click);
			this.btArticuloSRI.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btArticuloSRI_MouseDown);
			// 
			// label46
			// 
			this.label46.AutoSize = true;
			this.label46.BackColor = System.Drawing.Color.Transparent;
			this.label46.Location = new System.Drawing.Point(11, 137);
			this.label46.Name = "label46";
			this.label46.Size = new System.Drawing.Size(33, 16);
			this.label46.TabIndex = 116;
			this.label46.Text = "Renta";
			// 
			// txtCodigo
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance1;
			this.txtCodigo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCodigo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsArticulo, "Articulo.Codigo"));
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Location = new System.Drawing.Point(106, 39);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(270, 22);
			this.txtCodigo.TabIndex = 2;
			this.txtCodigo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCodigo_KeyDown);
			this.txtCodigo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCodigo_KeyPress);
			this.txtCodigo.ValueChanged += new System.EventHandler(this.txtCodigo_ValueChanged);
			// 
			// txtNombre
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance2;
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsArticulo, "Articulo.Articulo"));
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Location = new System.Drawing.Point(106, 67);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(323, 22);
			this.txtNombre.TabIndex = 3;
			this.txtNombre.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNombre_KeyDown);
			this.txtNombre.ValueChanged += new System.EventHandler(this.txtNombre_ValueChanged);
			// 
			// cmbGrupo
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbGrupo.Appearance = appearance3;
			this.cmbGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbGrupo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsArticulo, "Articulo.idGrupoArticulo"));
			this.cmbGrupo.DataSource = this.cvdArticuloGrupo;
			appearance4.BackColor = System.Drawing.Color.White;
			this.cmbGrupo.DisplayLayout.Appearance = appearance4;
			ultraGridColumn1.Header.VisiblePosition = 6;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 7;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Width = 336;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 2;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 8;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 26;
			ultraGridColumn9.Header.VisiblePosition = 9;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 27;
			ultraGridColumn10.Header.VisiblePosition = 10;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 57;
			ultraGridColumn11.Header.VisiblePosition = 1;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 120;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 237;
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 17;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 18;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 19;
			ultraGridColumn20.Hidden = true;
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
			this.cmbGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance5.BackColor = System.Drawing.Color.Transparent;
			this.cmbGrupo.DisplayLayout.Override.CardAreaAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 10F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.cmbGrupo.DisplayLayout.Override.HeaderAppearance = appearance6;
			this.cmbGrupo.DisplayLayout.Override.RowFilterAction = Infragistics.Win.UltraWinGrid.RowFilterAction.HideFilteredOutRows;
			this.cmbGrupo.DisplayLayout.Override.RowFilterMode = Infragistics.Win.UltraWinGrid.RowFilterMode.SiblingRowsOnly;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbGrupo.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			this.cmbGrupo.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbGrupo.DisplayLayout.Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.Fixed;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbGrupo.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.cmbGrupo.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.cmbGrupo.DisplayMember = "Grupo";
			this.cmbGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbGrupo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGrupo.Location = new System.Drawing.Point(106, 9);
			this.cmbGrupo.Name = "cmbGrupo";
			this.cmbGrupo.Size = new System.Drawing.Size(192, 21);
			this.cmbGrupo.TabIndex = 1;
			this.cmbGrupo.ValueMember = "idGrupoArticulo";
			this.cmbGrupo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbGrupo_KeyPress);
			this.cmbGrupo.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbGrupo_InitializeLayout);
			// 
			// cvdArticuloGrupo
			// 
			this.cvdArticuloGrupo.BindingContextCtrl = this;
			this.cvdArticuloGrupo.DataSet = this.cdsArticuloTabla;
			this.cvdArticuloGrupo.Sort = "Grupo";
			this.cvdArticuloGrupo.TableName = "";
			this.cvdArticuloGrupo.TableViewName = "ArticuloGrupo";
			// 
			// btGrupo
			// 
			this.btGrupo.Image = ((System.Drawing.Image)(resources.GetObject("btGrupo.Image")));
			this.btGrupo.Location = new System.Drawing.Point(306, 7);
			this.btGrupo.Name = "btGrupo";
			this.btGrupo.Size = new System.Drawing.Size(24, 25);
			this.btGrupo.TabIndex = 114;
			this.btGrupo.ToolTip = "Grupo de artículos";
			this.btGrupo.Click += new System.EventHandler(this.btGrupo_Click);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(10, 70);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(65, 16);
			this.label3.TabIndex = 113;
			this.label3.Text = "Descripción:";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(10, 11);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(39, 16);
			this.label2.TabIndex = 112;
			this.label2.Text = "Grupo:";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(10, 42);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(42, 16);
			this.label1.TabIndex = 111;
			this.label1.Text = "Código:";
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "Articulo";
			this.barraDato1.DataNombreId = "idArticulo";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsArticulo;
			this.barraDato1.DataTabla = "Articulo";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Location = new System.Drawing.Point(8, 210);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(256, 26);
			this.barraDato1.TabIndex = 0;
			this.barraDato1.VisibleBorrar = false;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = false;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Buscar += new System.EventHandler(this.barraDato1_Buscar);
			this.barraDato1.Load += new System.EventHandler(this.barraDato1_Load);
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// txtIdArticulo
			// 
			this.txtIdArticulo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIdArticulo.Appearance = appearance9;
			this.txtIdArticulo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsArticulo, "Articulo.idArticulo"));
			this.txtIdArticulo.Location = new System.Drawing.Point(344, 8);
			this.txtIdArticulo.Name = "txtIdArticulo";
			this.txtIdArticulo.Size = new System.Drawing.Size(100, 22);
			this.txtIdArticulo.TabIndex = 118;
			this.txtIdArticulo.Visible = false;
			// 
			// txtIva
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva.Appearance = appearance10;
			this.txtIva.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsArticulo, "Articulo.Iva"));
			this.txtIva.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva.Enabled = false;
			this.txtIva.FormatString = "#,##0.00";
			this.txtIva.Location = new System.Drawing.Point(344, 104);
			this.txtIva.Name = "txtIva";
			this.txtIva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva.PromptChar = ' ';
			this.txtIva.Size = new System.Drawing.Size(56, 22);
			this.txtIva.TabIndex = 4;
			this.txtIva.Visible = false;
			this.txtIva.Click += new System.EventHandler(this.txtIva_Click);
			this.txtIva.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtIva_KeyPress);
			this.txtIva.ValueChanged += new System.EventHandler(this.txtIva_ValueChanged);
			this.txtIva.Enter += new System.EventHandler(this.txtIva_Enter);
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Location = new System.Drawing.Point(304, 104);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(26, 16);
			this.label9.TabIndex = 120;
			this.label9.Text = "IVA:";
			this.label9.Visible = false;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(416, 104);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(15, 16);
			this.label4.TabIndex = 121;
			this.label4.Text = "%";
			this.label4.Visible = false;
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
			// btLimpia
			// 
			this.btLimpia.Location = new System.Drawing.Point(72, 133);
			this.btLimpia.Name = "btLimpia";
			this.btLimpia.Size = new System.Drawing.Size(24, 25);
			this.btLimpia.TabIndex = 122;
			this.btLimpia.Text = "-";
			this.btLimpia.ToolTip = "Retenciones de artículos";
			this.btLimpia.Click += new System.EventHandler(this.btLimpia_Click);
			// 
			// btnLimpia1
			// 
			this.btnLimpia1.Location = new System.Drawing.Point(72, 164);
			this.btnLimpia1.Name = "btnLimpia1";
			this.btnLimpia1.Size = new System.Drawing.Size(24, 25);
			this.btnLimpia1.TabIndex = 126;
			this.btnLimpia1.Text = "-";
			this.btnLimpia1.ToolTip = "Retenciones de artículos";
			this.btnLimpia1.Visible = false;
			this.btnLimpia1.Click += new System.EventHandler(this.btnLimpia1_Click);
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(8, 168);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(64, 16);
			this.label5.TabIndex = 125;
			this.label5.Text = "S.R.I. I.V.A.";
			this.label5.Visible = false;
			// 
			// cmbRetencionIVA
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbRetencionIVA.Appearance = appearance11;
			this.cmbRetencionIVA.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbRetencionIVA.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbRetencionIVA.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsArticulo, "Articulo.idRetIVA"));
			this.cmbRetencionIVA.DataSource = this.ultraDataSource1;
			ultraGridColumn21.Header.VisiblePosition = 0;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 1;
			ultraGridColumn22.Width = 100;
			ultraGridColumn23.Header.VisiblePosition = 2;
			ultraGridColumn23.Width = 250;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23});
			this.cmbRetencionIVA.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbRetencionIVA.DisplayMember = "Detalle";
			this.cmbRetencionIVA.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbRetencionIVA.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbRetencionIVA.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbRetencionIVA.Location = new System.Drawing.Point(107, 168);
			this.cmbRetencionIVA.Name = "cmbRetencionIVA";
			this.cmbRetencionIVA.Size = new System.Drawing.Size(285, 21);
			this.cmbRetencionIVA.TabIndex = 127;
			this.cmbRetencionIVA.ValueMember = "idRetIva";
			this.cmbRetencionIVA.Visible = false;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(System.Double);
			ultraDataColumn2.DefaultValue = 0;
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3});
			// 
			// cmbArticuloSRI
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbArticuloSRI.Appearance = appearance12;
			this.cmbArticuloSRI.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbArticuloSRI.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbArticuloSRI.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsArticulo, "Articulo.idArticuloSRI"));
			this.cmbArticuloSRI.DataSource = this.ultraDataSource2;
			ultraGridColumn24.Header.VisiblePosition = 0;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 1;
			ultraGridColumn25.Width = 75;
			ultraGridColumn26.Header.VisiblePosition = 2;
			ultraGridColumn26.Width = 350;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26});
			this.cmbArticuloSRI.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbArticuloSRI.DisplayMember = "Detalle";
			this.cmbArticuloSRI.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbArticuloSRI.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbArticuloSRI.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbArticuloSRI.Location = new System.Drawing.Point(107, 136);
			this.cmbArticuloSRI.Name = "cmbArticuloSRI";
			this.cmbArticuloSRI.Size = new System.Drawing.Size(285, 21);
			this.cmbArticuloSRI.TabIndex = 128;
			this.cmbArticuloSRI.ValueMember = "idArticuloSRI";
			// 
			// ultraDataSource2
			// 
			ultraDataColumn4.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			// 
			// chkImpuestoIVA
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkImpuestoIVA.Appearance = appearance13;
			this.chkImpuestoIVA.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsArticulo, "Articulo.ImpuestoIVA"));
			this.chkImpuestoIVA.Enabled = false;
			this.chkImpuestoIVA.Location = new System.Drawing.Point(104, 100);
			this.chkImpuestoIVA.Name = "chkImpuestoIVA";
			this.chkImpuestoIVA.Size = new System.Drawing.Size(104, 22);
			this.chkImpuestoIVA.TabIndex = 129;
			this.chkImpuestoIVA.Text = "Impuesto I.V.A.";
			// 
			// frmArticulosServicios
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(448, 240);
			this.Controls.Add(this.chkImpuestoIVA);
			this.Controls.Add(this.cmbArticuloSRI);
			this.Controls.Add(this.cmbRetencionIVA);
			this.Controls.Add(this.btnLimpia1);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label46);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.btLimpia);
			this.Controls.Add(this.txtIva);
			this.Controls.Add(this.txtIdArticulo);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.btArticuloSRI);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.cmbGrupo);
			this.Controls.Add(this.btGrupo);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.MaximizeBox = false;
			this.Name = "frmArticulosServicios";
			this.Text = "Articulos de Servicios";
			this.Load += new System.EventHandler(this.frmArticulosServicios_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticuloTabla)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cvdArticuloGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRetencionIVA)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticuloSRI)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		C1.Data.C1DataRow drArticulo;

		private void cdsArticulo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsArticulo.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsArticuloTabla_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsArticuloTabla.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void barraDato1_Load(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			this.barraDato1.BorraRegistro();
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
//			if (!miAcceso.Nuevo)
//			{
//				MessageBox.Show("No tiene acceso a Crear Nuevos Servicios.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
//				return;
//			}

			this.barraDato1.CrearRegistro();
			this.txtCodigo.Enabled = false;
//			this.txtNombre.Enabled = false;
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			this.barraDato1.DeshacerRegistro();
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			if (drArticulo == null)
			{
				MessageBox.Show("No existe un registro para editar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
//			if (!miAcceso.Editar)
//			{
//				MessageBox.Show("No tiene acceso a Editar Articulos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
//				return;
//			}

			this.barraDato1.EditarRegistro();
//			this.txtNombre.Enabled = false;

			this.txtCodigo.Enabled = false;
			if (!miAcceso.BModificaRenta) this.cmbArticuloSRI.Enabled = false;
			if (!miAcceso.BModificaIVA) this.cmbRetencionIVA.Enabled = false;
			
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			#region Validacion
			if (drArticulo["idGrupoArticulo"] == DBNull.Value)
			{
				MessageBox.Show("Seleccione un Grupo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbGrupo.Focus();
				return;
			}
			if (drArticulo["Codigo"] == DBNull.Value)
			{
				MessageBox.Show("Escriba el código", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCodigo.Focus();
				return;
			}
			if (drArticulo["Articulo"] == DBNull.Value)
			{
				MessageBox.Show("Escriba la descripción del artículo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNombre.Focus();
				return;
			}
			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From Articulo Where Codigo = '{0}' And idTipoGrupo Not In (1, 6) And idArticulo <> {1}", this.txtCodigo.Text.ToString(), (int)drArticulo["idArticulo"])) > 1)
			{
				MessageBox.Show("Este Codigo ya existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNombre.Focus();
				return;
			}
			#endregion Validacion

			drArticulo["Origen"] = 1;
			
			drArticulo["idTipoGrupo"] = Funcion.iEscalarSQL(cdsArticulo, string.Format("Select IsNull(idTipoGrupo, 0) From ArticuloGrupo Where idGrupoArticulo = {0}", (int)drArticulo["idGrupoArticulo"]));
			int ada = (int)drArticulo["idTipoGrupo"];
			this.barraDato1.GrabaRegistro();
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
			try
			{
				drArticulo = cdsArticulo.TableViews["Articulo"].Rows[0];
			}
			catch{};
		}

		private void Busqueda(int idBusca)
		{
			barraDato1.IdRegistro = idBusca;
			barraDato1.ProximoId(5);
		}
		private void UnloadMe()
		{
			this.Close();
		}
		private void frmArticulosServicios_Load(object sender, System.EventArgs e)
		{
			if (!Funcion.Permiso("54",cdsSeteoF))
			{
				MessageBox.Show("No puede ingresar a Servicios", "Point Technology");
				BeginInvoke(new MethodInvoker(UnloadMe));
			}

			miAcceso = new Acceso(cdsSeteoF, "0412");

			barraDato1.stFiltroExt = " And idTipogrupo Not In (1, 6)";

			this.cmbArticuloSRI.DataSource = Funcion.dvProcedimiento (cdsSeteoF, "Select idArticuloSRI, Codigo, Detalle From ArticuloSRI Where Activo = 1 Order By Codigo");

			#region Busqueda
			if (idArticulo > 0)
				Busqueda(idArticulo);
			else
				barraDato1.PosUltima();

			this.barraDato1.HabilitaControles(false);
			#endregion Busqueda

//			this.cmbRetencionIVA.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idRetIva, Porcentaje, Detalle From RetencionIva Order by Porcentaje");

			string sFiltro = "";
			sFiltro = "idTipoGrupo Not In (1, 6)";
						
			C1.Data.FilterCondition fcP = new C1.Data.FilterCondition(cdsArticuloTabla, "ArticuloGrupo", sFiltro);
			C1.Data.FilterConditions fcT = new C1.Data.FilterConditions();
			fcT.Add(fcP);
			cdsArticuloTabla.Fill(fcT, false);
		}

		private void btArticuloSRI_Click(object sender, System.EventArgs e)
		{
			using(ArticuloSRI miArt = new ArticuloSRI())
			{
				miArt.ShowDialog();
			}
			cdsArticuloTabla.Clear();
			cdsArticuloTabla.Fill();
			cmbArticuloSRI.Refresh();
		}

		private void btGrupo_Click(object sender, System.EventArgs e)
		{
			ArtGrupo miGrupo;
			if (cmbGrupo.Value != DBNull.Value && cmbGrupo.Value != null)
				miGrupo = new ArtGrupo((int) cmbGrupo.Value, 2);
			else
				miGrupo = new ArtGrupo();
			miGrupo.ShowDialog();
			miGrupo.Dispose();
			miGrupo = null;
			cdsArticuloTabla.Clear();
			cdsArticuloTabla.Fill();
			cmbGrupo.Refresh();
		}

		private void txtCodigo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			
		}

		private void txtIva_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtIva_Click(object sender, System.EventArgs e)
		{
			this.txtIva.SelectAll();
		}

		private void txtIva_Enter(object sender, System.EventArgs e)
		{
			this.txtIva.SelectAll();
		}

		private void txtCodigo_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cmbGrupo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtNombre.Focus();
		}

		private void txtCodigo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtIva.Focus();
		}

		private void txtIva_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.cmbArticuloSRI.Focus();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
			if (drArticulo == null)
			{
				MessageBox.Show("No existe un registros de servicios.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
//			if (!miAcceso.Buscar)
//			{
//				MessageBox.Show("No tiene acceso a buscar servicios.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
//				return;
//			}

			using (BuscaBienesServicios miBusqueda = new BuscaBienesServicios("", 0))
			{
				if (DialogResult.OK == miBusqueda.ShowDialog())
				{					
					barraDato1.IdRegistro = miBusqueda.idArticulo;
					barraDato1.ProximoId(5);					
				}
			}
		}

		private void btArticuloSRI_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (!this.barraDato1.bEditar)return;
			if (e.Button != MouseButtons.Right) return;
			drArticulo["idArticuloSRI"] = System.DBNull.Value;
		}

		private void btLimpia_Click(object sender, System.EventArgs e)
		{
			if (!this.barraDato1.bEditar) return;
			this.cmbArticuloSRI.Value = System.DBNull.Value;
		}

		private void btnLimpia1_Click(object sender, System.EventArgs e)
		{
			if (!this.barraDato1.bEditar) return;
			this.cmbRetencionIVA.Value = System.DBNull.Value;
		}

		private void cmbRetencionIVA_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.cmbRetencionIVA.Focus();
		}

		private void txtNombre_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{
				using (frmBuscaCuentas miBusqueda = new frmBuscaCuentas())
				{
					if (DialogResult.OK == miBusqueda.ShowDialog())
					{
						drArticulo["Codigo"] = miBusqueda.Cuenta;
						drArticulo["Articulo"] = miBusqueda.ultraGrid1.ActiveRow.Cells["Descripcion"].Value.ToString();
					}
				}
			}
		}

		private void cmbGrupo_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtNombre_ValueChanged(object sender, System.EventArgs e)
		{
		
		}
	}
}

