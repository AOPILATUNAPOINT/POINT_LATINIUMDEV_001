using System.Data;
using System.Data.SqlClient;
using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCre_Garante.
	/// </summary>
	public class frmCre_Garante : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtpNEFechaNacimiento;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoRuc;
		private System.Windows.Forms.Label label1;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtDocumento;
		private System.Windows.Forms.Label lblBodega;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmdCalificacion;
		private System.Windows.Forms.Label label18;
		private System.Windows.Forms.Label label24;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEstado;
		private System.Windows.Forms.Label label22;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSnombre;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPnombre;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtApellidoM;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtApellidoP;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNEBarrio;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNEParroquia;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNECanton;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNEProvincia;
		private System.Windows.Forms.Label label19;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label11;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNETelefono;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCalle2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumcasa;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCalle1;
		private System.Windows.Forms.Label label14;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private System.Windows.Forms.Label label15;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas2;
		private System.Windows.Forms.Button btnGuardar;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource8;
		private System.Windows.Forms.Button btnCancelar;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCre_Solicitud;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidGarante;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtScore;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidIdentificacion;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCre_DatosUbicacion;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCre_DatosGenerales;
    string sCreSolicitud ;
		int idCre_Solicitud = 0;
		bool bCarga ;
		bool bgarante = true;
		bool bPantallaCG = true;
		int isolicitud = 0;
		bool bnuevo = true;

		public frmCre_Garante(string screSolicitud, int IdCre_Solicitud, bool bGarante, bool BPantallaCG, int iSolicitud, bool bNuevo)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			sCreSolicitud = screSolicitud;
			idCre_Solicitud = IdCre_Solicitud;
			bgarante = bGarante;
			bPantallaCG = BPantallaCG;
			isolicitud = iSolicitud;
			bnuevo = bNuevo;

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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCre_Garante));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCalificacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoCalificacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBarrio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBarrio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idParroquia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idParroquia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCanton");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCanton");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProvincia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNegadoPendiente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			this.dtpNEFechaNacimiento = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbTipoRuc = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.txtDocumento = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblBodega = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.cmdCalificacion = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label18 = new System.Windows.Forms.Label();
			this.label24 = new System.Windows.Forms.Label();
			this.txtScore = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbEstado = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label22 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.txtSnombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtPnombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtApellidoM = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtApellidoP = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label10 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.cmbNEBarrio = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbNEParroquia = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbNECanton = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbNEProvincia = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label19 = new System.Windows.Forms.Label();
			this.label13 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.txtNETelefono = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCalle2 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNumcasa = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCalle1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label14 = new System.Windows.Forms.Label();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label15 = new System.Windows.Forms.Label();
			this.txtNotas2 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource8 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.txtIdCre_Solicitud = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidGarante = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidIdentificacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdCre_DatosUbicacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdCre_DatosGenerales = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			((System.ComponentModel.ISupportInitialize)(this.dtpNEFechaNacimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDocumento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmdCalificacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtScore)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSnombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPnombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtApellidoM)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtApellidoP)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEBarrio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEParroquia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNECanton)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEProvincia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNETelefono)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCalle2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumcasa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCalle1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCre_Solicitud)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidGarante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidIdentificacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCre_DatosUbicacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCre_DatosGenerales)).BeginInit();
			this.SuspendLayout();
			// 
			// dtpNEFechaNacimiento
			// 
			this.dtpNEFechaNacimiento.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtpNEFechaNacimiento.DateButtons.Add(dateButton1);
			this.dtpNEFechaNacimiento.Enabled = false;
			this.dtpNEFechaNacimiento.Format = "dd/MM/yyyy";
			this.dtpNEFechaNacimiento.Location = new System.Drawing.Point(56, 14);
			this.dtpNEFechaNacimiento.Name = "dtpNEFechaNacimiento";
			this.dtpNEFechaNacimiento.NonAutoSizeHeight = 23;
			this.dtpNEFechaNacimiento.Size = new System.Drawing.Size(104, 21);
			this.dtpNEFechaNacimiento.TabIndex = 664;
			this.dtpNEFechaNacimiento.Value = ((object)(resources.GetObject("dtpNEFechaNacimiento.Value")));
			// 
			// cmbTipoRuc
			// 
			this.cmbTipoRuc.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoRuc.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoRuc.DisplayMember = "Nombre";
			this.cmbTipoRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoRuc.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoRuc.Enabled = false;
			this.cmbTipoRuc.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoRuc.Location = new System.Drawing.Point(272, 13);
			this.cmbTipoRuc.Name = "cmbTipoRuc";
			this.cmbTipoRuc.Size = new System.Drawing.Size(144, 22);
			this.cmbTipoRuc.TabIndex = 774;
			this.cmbTipoRuc.ValueMember = "idTipoDoc";
			this.cmbTipoRuc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbTipoRuc_KeyPress);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(424, 16);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(62, 16);
			this.label1.TabIndex = 775;
			this.label1.Text = "Documento";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDocumento
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDocumento.Appearance = appearance1;
			this.txtDocumento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDocumento.Location = new System.Drawing.Point(496, 14);
			this.txtDocumento.MaxLength = 10;
			this.txtDocumento.Name = "txtDocumento";
			this.txtDocumento.Size = new System.Drawing.Size(136, 21);
			this.txtDocumento.TabIndex = 776;
			this.txtDocumento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDocumento_KeyDown);
			this.txtDocumento.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDocumento_KeyPress);
			this.txtDocumento.Validating += new System.ComponentModel.CancelEventHandler(this.txtDocumento_Validating);
			this.txtDocumento.Validated += new System.EventHandler(this.txtDocumento_Validated);
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(176, 16);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(87, 16);
			this.lblBodega.TabIndex = 777;
			this.lblBodega.Text = "Tipo Documento";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 16);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(36, 16);
			this.label2.TabIndex = 778;
			this.label2.Text = "Fecha";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmdCalificacion
			// 
			this.cmdCalificacion.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmdCalificacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmdCalificacion.DataSource = this.ultraDataSource6;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.Caption = "Calificacion";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmdCalificacion.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmdCalificacion.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmdCalificacion.DisplayMember = "Codigo";
			this.cmdCalificacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmdCalificacion.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmdCalificacion.Enabled = false;
			this.cmdCalificacion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmdCalificacion.Location = new System.Drawing.Point(80, 45);
			this.cmdCalificacion.Name = "cmdCalificacion";
			this.cmdCalificacion.Size = new System.Drawing.Size(112, 22);
			this.cmdCalificacion.TabIndex = 783;
			this.cmdCalificacion.ValueMember = "idTipoCalificacion";
			this.cmdCalificacion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmdCalificacion_KeyDown);
			// 
			// ultraDataSource6
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(8, 48);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(62, 16);
			this.label18.TabIndex = 782;
			this.label18.Text = "Calificación";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label24
			// 
			this.label24.AutoSize = true;
			this.label24.Location = new System.Drawing.Point(208, 48);
			this.label24.Name = "label24";
			this.label24.Size = new System.Drawing.Size(34, 16);
			this.label24.TabIndex = 785;
			this.label24.Text = "Score";
			this.label24.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtScore
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtScore.Appearance = appearance2;
			this.txtScore.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtScore.Enabled = false;
			this.txtScore.FormatString = "0";
			this.txtScore.Location = new System.Drawing.Point(272, 46);
			this.txtScore.MinValue = 0;
			this.txtScore.Name = "txtScore";
			this.txtScore.PromptChar = ' ';
			this.txtScore.Size = new System.Drawing.Size(80, 21);
			this.txtScore.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtScore.TabIndex = 784;
			this.txtScore.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtScore_KeyDown);
			// 
			// cmbEstado
			// 
			this.cmbEstado.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEstado.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEstado.DataSource = this.ultraDataSource7;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbEstado.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbEstado.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbEstado.DisplayMember = "Estado";
			this.cmbEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEstado.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEstado.Enabled = false;
			this.cmbEstado.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstado.Location = new System.Drawing.Point(496, 45);
			this.cmbEstado.Name = "cmbEstado";
			this.cmbEstado.Size = new System.Drawing.Size(120, 22);
			this.cmbEstado.TabIndex = 787;
			this.cmbEstado.ValueMember = "idEstado";
			this.cmbEstado.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbEstado_KeyDown);
			// 
			// ultraDataSource7
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// label22
			// 
			this.label22.AutoSize = true;
			this.label22.Location = new System.Drawing.Point(424, 48);
			this.label22.Name = "label22";
			this.label22.Size = new System.Drawing.Size(39, 16);
			this.label22.TabIndex = 786;
			this.label22.Text = "Estado";
			this.label22.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(312, 122);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(93, 16);
			this.label5.TabIndex = 791;
			this.label5.Text = "Segundo Nombre";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(8, 122);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(81, 16);
			this.label4.TabIndex = 790;
			this.label4.Text = "Primer Nombre";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(312, 90);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(89, 16);
			this.label3.TabIndex = 789;
			this.label3.Text = "Apellido Materno";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 90);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(87, 16);
			this.label6.TabIndex = 788;
			this.label6.Text = "Apellido Paterno";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSnombre
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSnombre.Appearance = appearance3;
			this.txtSnombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtSnombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSnombre.Enabled = false;
			this.txtSnombre.Location = new System.Drawing.Point(432, 120);
			this.txtSnombre.Name = "txtSnombre";
			this.txtSnombre.Size = new System.Drawing.Size(184, 21);
			this.txtSnombre.TabIndex = 795;
			this.txtSnombre.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSnombre_KeyDown);
			// 
			// txtPnombre
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPnombre.Appearance = appearance4;
			this.txtPnombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtPnombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPnombre.Enabled = false;
			this.txtPnombre.Location = new System.Drawing.Point(104, 120);
			this.txtPnombre.Name = "txtPnombre";
			this.txtPnombre.Size = new System.Drawing.Size(184, 21);
			this.txtPnombre.TabIndex = 794;
			this.txtPnombre.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtPnombre_KeyDown);
			// 
			// txtApellidoM
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtApellidoM.Appearance = appearance5;
			this.txtApellidoM.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtApellidoM.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtApellidoM.Enabled = false;
			this.txtApellidoM.Location = new System.Drawing.Point(432, 88);
			this.txtApellidoM.Name = "txtApellidoM";
			this.txtApellidoM.Size = new System.Drawing.Size(184, 21);
			this.txtApellidoM.TabIndex = 793;
			this.txtApellidoM.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtApellidoM_KeyDown);
			// 
			// txtApellidoP
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtApellidoP.Appearance = appearance6;
			this.txtApellidoP.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtApellidoP.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtApellidoP.Enabled = false;
			this.txtApellidoP.Location = new System.Drawing.Point(104, 88);
			this.txtApellidoP.Name = "txtApellidoP";
			this.txtApellidoP.Size = new System.Drawing.Size(184, 21);
			this.txtApellidoP.TabIndex = 792;
			this.txtApellidoP.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtApellidoP_KeyDown);
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.SystemColors.Control;
			this.label10.Location = new System.Drawing.Point(320, 192);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(34, 16);
			this.label10.TabIndex = 802;
			this.label10.Text = "Barrio";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.SystemColors.Control;
			this.label9.Location = new System.Drawing.Point(8, 192);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(53, 16);
			this.label9.TabIndex = 800;
			this.label9.Text = "Parroquia";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.SystemColors.Control;
			this.label8.Location = new System.Drawing.Point(320, 155);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(41, 16);
			this.label8.TabIndex = 798;
			this.label8.Text = "Canton";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.SystemColors.Control;
			this.label7.Location = new System.Drawing.Point(8, 155);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(51, 16);
			this.label7.TabIndex = 796;
			this.label7.Text = "Provincia";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbNEBarrio
			// 
			this.cmbNEBarrio.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNEBarrio.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNEBarrio.DataSource = this.ultraDataSource4;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.Caption = "Barrio";
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 170;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbNEBarrio.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbNEBarrio.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbNEBarrio.DisplayMember = "Nombre";
			this.cmbNEBarrio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNEBarrio.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNEBarrio.Enabled = false;
			this.cmbNEBarrio.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNEBarrio.Location = new System.Drawing.Point(432, 189);
			this.cmbNEBarrio.Name = "cmbNEBarrio";
			this.cmbNEBarrio.Size = new System.Drawing.Size(184, 22);
			this.cmbNEBarrio.TabIndex = 803;
			this.cmbNEBarrio.ValueMember = "idBarrio";
			this.cmbNEBarrio.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNEBarrio_KeyDown);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn5.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			// 
			// cmbNEParroquia
			// 
			this.cmbNEParroquia.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNEParroquia.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNEParroquia.DataSource = this.ultraDataSource3;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.Caption = "Provincia";
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Width = 160;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.cmbNEParroquia.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbNEParroquia.DisplayMember = "Nombre";
			this.cmbNEParroquia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNEParroquia.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNEParroquia.Enabled = false;
			this.cmbNEParroquia.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNEParroquia.Location = new System.Drawing.Point(104, 189);
			this.cmbNEParroquia.Name = "cmbNEParroquia";
			this.cmbNEParroquia.Size = new System.Drawing.Size(184, 22);
			this.cmbNEParroquia.TabIndex = 801;
			this.cmbNEParroquia.ValueMember = "idParroquia";
			this.cmbNEParroquia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNEParroquia_KeyDown);
			this.cmbNEParroquia.ValueChanged += new System.EventHandler(this.cmbNEParroquia_ValueChanged);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn7.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn7,
																																 ultraDataColumn8});
			// 
			// cmbNECanton
			// 
			this.cmbNECanton.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNECanton.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNECanton.DataSource = this.ultraDataSource2;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.Caption = "Canton";
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Width = 160;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			this.cmbNECanton.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbNECanton.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbNECanton.DisplayMember = "Nombre";
			this.cmbNECanton.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNECanton.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNECanton.Enabled = false;
			this.cmbNECanton.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNECanton.Location = new System.Drawing.Point(432, 152);
			this.cmbNECanton.Name = "cmbNECanton";
			this.cmbNECanton.Size = new System.Drawing.Size(184, 22);
			this.cmbNECanton.TabIndex = 799;
			this.cmbNECanton.ValueMember = "idCanton";
			this.cmbNECanton.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNECanton_KeyDown);
			this.cmbNECanton.ValueChanged += new System.EventHandler(this.cmbNECanton_ValueChanged);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn9.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn9,
																																 ultraDataColumn10});
			// 
			// cmbNEProvincia
			// 
			this.cmbNEProvincia.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNEProvincia.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNEProvincia.DataSource = this.ultraDataSource1;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 50;
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Width = 170;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.cmbNEProvincia.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbNEProvincia.DisplayMember = "Nombre";
			this.cmbNEProvincia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNEProvincia.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNEProvincia.Enabled = false;
			this.cmbNEProvincia.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNEProvincia.Location = new System.Drawing.Point(104, 152);
			this.cmbNEProvincia.Name = "cmbNEProvincia";
			this.cmbNEProvincia.Size = new System.Drawing.Size(184, 22);
			this.cmbNEProvincia.TabIndex = 797;
			this.cmbNEProvincia.ValueMember = "idProvincia";
			this.cmbNEProvincia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNEProvincia_KeyDown);
			this.cmbNEProvincia.ValueChanged += new System.EventHandler(this.cmbNEProvincia_ValueChanged);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn11.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn11,
																																 ultraDataColumn12});
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.BackColor = System.Drawing.SystemColors.Control;
			this.label19.Location = new System.Drawing.Point(408, 266);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(48, 16);
			this.label19.TabIndex = 810;
			this.label19.Text = "Teléfono";
			this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.SystemColors.Control;
			this.label13.Location = new System.Drawing.Point(8, 266);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(91, 16);
			this.label13.TabIndex = 808;
			this.label13.Text = "Calle Secundaria";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.SystemColors.Control;
			this.label12.Location = new System.Drawing.Point(408, 232);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(90, 16);
			this.label12.TabIndex = 806;
			this.label12.Text = "Número de Casa";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.BackColor = System.Drawing.SystemColors.Control;
			this.label11.Location = new System.Drawing.Point(8, 232);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(77, 16);
			this.label11.TabIndex = 804;
			this.label11.Text = "Calle Principal";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNETelefono
			// 
			this.txtNETelefono.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNETelefono.Enabled = false;
			this.txtNETelefono.Location = new System.Drawing.Point(512, 264);
			this.txtNETelefono.MaxLength = 9;
			this.txtNETelefono.Name = "txtNETelefono";
			this.txtNETelefono.Size = new System.Drawing.Size(104, 21);
			this.txtNETelefono.TabIndex = 811;
			this.txtNETelefono.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNETelefono_KeyDown);
			this.txtNETelefono.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNETelefono_KeyPress);
			// 
			// txtCalle2
			// 
			this.txtCalle2.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCalle2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCalle2.Enabled = false;
			this.txtCalle2.Location = new System.Drawing.Point(104, 264);
			this.txtCalle2.Name = "txtCalle2";
			this.txtCalle2.Size = new System.Drawing.Size(272, 21);
			this.txtCalle2.TabIndex = 809;
			this.txtCalle2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCalle2_KeyDown);
			// 
			// txtNumcasa
			// 
			this.txtNumcasa.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNumcasa.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumcasa.Enabled = false;
			this.txtNumcasa.Location = new System.Drawing.Point(512, 232);
			this.txtNumcasa.Name = "txtNumcasa";
			this.txtNumcasa.Size = new System.Drawing.Size(104, 21);
			this.txtNumcasa.TabIndex = 807;
			this.txtNumcasa.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNumcasa_KeyDown);
			// 
			// txtCalle1
			// 
			this.txtCalle1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCalle1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCalle1.Enabled = false;
			this.txtCalle1.Location = new System.Drawing.Point(104, 230);
			this.txtCalle1.Name = "txtCalle1";
			this.txtCalle1.Size = new System.Drawing.Size(272, 21);
			this.txtCalle1.TabIndex = 805;
			this.txtCalle1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCalle1_KeyDown);
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.SystemColors.Control;
			this.label14.Location = new System.Drawing.Point(8, 296);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(128, 16);
			this.label14.TabIndex = 813;
			this.label14.Text = "Referencia de Ubicación";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNotas
			// 
			this.txtNotas.AcceptsReturn = true;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance7;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Enabled = false;
			this.txtNotas.Location = new System.Drawing.Point(8, 320);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(304, 48);
			this.txtNotas.TabIndex = 812;
			this.txtNotas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNotas_KeyDown);
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.BackColor = System.Drawing.SystemColors.Control;
			this.label15.Location = new System.Drawing.Point(320, 296);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(80, 16);
			this.label15.TabIndex = 815;
			this.label15.Text = "Observaciones";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNotas2
			// 
			this.txtNotas2.AcceptsReturn = true;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas2.Appearance = appearance8;
			this.txtNotas2.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas2.Enabled = false;
			this.txtNotas2.Location = new System.Drawing.Point(320, 320);
			this.txtNotas2.Multiline = true;
			this.txtNotas2.Name = "txtNotas2";
			this.txtNotas2.Size = new System.Drawing.Size(304, 48);
			this.txtNotas2.TabIndex = 814;
			this.txtNotas2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNotas2_KeyDown);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(8, 376);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(72, 23);
			this.btnGuardar.TabIndex = 816;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
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
			// ultraDataSource5
			// 
			ultraDataColumn13.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn13,
																																 ultraDataColumn14});
			// 
			// ultraDataSource8
			// 
			ultraDataColumn15.DataType = typeof(int);
			this.ultraDataSource8.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn15,
																																 ultraDataColumn16});
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnCancelar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(88, 376);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 817;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// txtIdCre_Solicitud
			// 
			this.txtIdCre_Solicitud.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtIdCre_Solicitud.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCre_Solicitud.Enabled = false;
			this.txtIdCre_Solicitud.Location = new System.Drawing.Point(400, 376);
			this.txtIdCre_Solicitud.Name = "txtIdCre_Solicitud";
			this.txtIdCre_Solicitud.PromptChar = ' ';
			this.txtIdCre_Solicitud.Size = new System.Drawing.Size(72, 21);
			this.txtIdCre_Solicitud.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtIdCre_Solicitud.TabIndex = 818;
			// 
			// txtidGarante
			// 
			this.txtidGarante.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtidGarante.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidGarante.Enabled = false;
			this.txtidGarante.Location = new System.Drawing.Point(480, 376);
			this.txtidGarante.Name = "txtidGarante";
			this.txtidGarante.PromptChar = ' ';
			this.txtidGarante.Size = new System.Drawing.Size(72, 21);
			this.txtidGarante.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidGarante.TabIndex = 819;
			// 
			// txtidIdentificacion
			// 
			this.txtidIdentificacion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtidIdentificacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidIdentificacion.Enabled = false;
			this.txtidIdentificacion.Location = new System.Drawing.Point(320, 376);
			this.txtidIdentificacion.Name = "txtidIdentificacion";
			this.txtidIdentificacion.PromptChar = ' ';
			this.txtidIdentificacion.Size = new System.Drawing.Size(72, 21);
			this.txtidIdentificacion.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidIdentificacion.TabIndex = 820;
			// 
			// txtIdCre_DatosUbicacion
			// 
			this.txtIdCre_DatosUbicacion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtIdCre_DatosUbicacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCre_DatosUbicacion.Enabled = false;
			this.txtIdCre_DatosUbicacion.Location = new System.Drawing.Point(176, 376);
			this.txtIdCre_DatosUbicacion.Name = "txtIdCre_DatosUbicacion";
			this.txtIdCre_DatosUbicacion.PromptChar = ' ';
			this.txtIdCre_DatosUbicacion.Size = new System.Drawing.Size(64, 21);
			this.txtIdCre_DatosUbicacion.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtIdCre_DatosUbicacion.TabIndex = 823;
			// 
			// txtIdCre_DatosGenerales
			// 
			this.txtIdCre_DatosGenerales.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtIdCre_DatosGenerales.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCre_DatosGenerales.Enabled = false;
			this.txtIdCre_DatosGenerales.Location = new System.Drawing.Point(248, 376);
			this.txtIdCre_DatosGenerales.Name = "txtIdCre_DatosGenerales";
			this.txtIdCre_DatosGenerales.PromptChar = ' ';
			this.txtIdCre_DatosGenerales.Size = new System.Drawing.Size(64, 21);
			this.txtIdCre_DatosGenerales.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtIdCre_DatosGenerales.TabIndex = 821;
			// 
			// frmCre_Garante
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(640, 414);
			this.Controls.Add(this.txtIdCre_DatosUbicacion);
			this.Controls.Add(this.txtIdCre_DatosGenerales);
			this.Controls.Add(this.txtidIdentificacion);
			this.Controls.Add(this.txtidGarante);
			this.Controls.Add(this.txtIdCre_Solicitud);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.txtNotas2);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.label19);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.txtNETelefono);
			this.Controls.Add(this.txtCalle2);
			this.Controls.Add(this.txtNumcasa);
			this.Controls.Add(this.txtCalle1);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.cmbNEBarrio);
			this.Controls.Add(this.cmbNEParroquia);
			this.Controls.Add(this.cmbNECanton);
			this.Controls.Add(this.cmbNEProvincia);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.txtSnombre);
			this.Controls.Add(this.txtPnombre);
			this.Controls.Add(this.txtApellidoM);
			this.Controls.Add(this.txtApellidoP);
			this.Controls.Add(this.cmbEstado);
			this.Controls.Add(this.label22);
			this.Controls.Add(this.label24);
			this.Controls.Add(this.txtScore);
			this.Controls.Add(this.cmdCalificacion);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtDocumento);
			this.Controls.Add(this.cmbTipoRuc);
			this.Controls.Add(this.dtpNEFechaNacimiento);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmCre_Garante";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Garantes";
			this.Load += new System.EventHandler(this.frmCre_Garante_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtpNEFechaNacimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDocumento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmdCalificacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtScore)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSnombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPnombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtApellidoM)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtApellidoP)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEBarrio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEParroquia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNECanton)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEProvincia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNETelefono)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCalle2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumcasa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCalle1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCre_Solicitud)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidGarante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidIdentificacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCre_DatosUbicacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCre_DatosGenerales)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmCre_Garante_Load(object sender, System.EventArgs e)
		{
			
			bCarga = true;
			this.cmbTipoRuc.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 4, 0, 0");
			this.cmbNEProvincia.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 0, 0, 0");
			this.cmdCalificacion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 8, 0, 0");
			this.cmbEstado.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 9, 0, 0");
			this.dtpNEFechaNacimiento.Value = DateTime.Today;
			this.cmbTipoRuc.Value = 1;
			this.txtIdCre_Solicitud.Value = idCre_Solicitud;
			HabilitaControles ();

			
			if (bPantallaCG)
			{
				if (idCre_Solicitud > 0)
				{
				
					string ssqlDo= string.Format ("Exec Cre_ConsultaGarantes {0}", isolicitud); 
					SqlDataReader drDO = Funcion.rEscalarSQL(cdsSeteoF,ssqlDo,true);
					drDO.Read();
					if(drDO.HasRows)
					{

						this.txtidGarante.Value = drDO.GetInt32 (0);
						this.txtIdCre_Solicitud.Value = drDO.GetInt32 (1);
						this.txtNotas2.Text = drDO.GetString (2);
					}
					drDO.Close();
			
					Consulta((int) this.txtidGarante.Value, (int) this.txtIdCre_Solicitud.Value);

				}
			}
			else
			{
			 Consulta(isolicitud, idCre_Solicitud);
			}

			 if (!bgarante ) Desabilitacontroles ();

				 bCarga = false;
		}
		private void HabilitaControles ()
		{
			#region Habilita 

			this.txtApellidoP.Enabled = true;
			this.txtApellidoM.Enabled= true;
			this.txtPnombre.Enabled = true;
			this.txtSnombre.Enabled = true; 
			this.cmbNEProvincia.Enabled = true;
			this.cmbNECanton.Enabled  = true;
			this.cmbNEParroquia.Enabled = true;
			this.cmbNEBarrio.Enabled = true;
			this.txtCalle1.Enabled = true;
			this.txtNumcasa.Enabled = true;
			this.txtCalle2.Enabled = true;
			this.txtNETelefono.Enabled = true;
			this.txtNotas.Enabled  = true;
			this.txtNotas2.Enabled = true;
			this.cmdCalificacion.Enabled  = true;
			this.cmbEstado.Enabled = true;
			this.txtScore.Enabled = true;
			this.btnGuardar.Enabled = true;

			#endregion Habilita 
		}
  
		private void LimpiaControles ()
		{
			#region Limpia 

			this.txtIdCre_Solicitud.Value = 0;
			this.txtidGarante.Value = 0;
			this.txtScore.Value = 0;
			this.txtApellidoP.Text ="";
			this.txtApellidoM.Text ="";
			this.txtPnombre.Text ="";
			this.txtSnombre.Text ="";
			this.dtpNEFechaNacimiento.Value = System.DBNull.Value;
			this.cmbNEProvincia.Value = System.DBNull.Value;
			this.cmbNECanton.Value = System.DBNull.Value;
			this.cmbNEParroquia.Value = System.DBNull.Value;
			this.cmbNEBarrio.Value = System.DBNull.Value;
			this.txtCalle1.Text ="";
			this.txtNumcasa.Text ="";
			this.txtCalle2.Text ="";
			this.txtNETelefono.Text ="";
			this.txtNotas.Text ="";
			this.txtNotas2.Text ="";
			this.cmdCalificacion.Value = System.DBNull.Value;
			this.cmbEstado.Value = System.DBNull.Value;
			this.txtidIdentificacion.Value = 0;
			this.txtDocumento.Text = "";
			this.cmbTipoRuc.Value = 0;
	
			#endregion Limpia 
		}

		private void Desabilitacontroles ()
		{
			#region Desabilitacontroles 


			this.txtApellidoP.Enabled = false;
			this.txtApellidoM.Enabled= false;
			this.txtPnombre.Enabled = false;;
			this.txtSnombre.Enabled = false; 
			this.dtpNEFechaNacimiento.Enabled = false;
			this.cmbNEProvincia.Enabled = false;
			this.cmbNECanton.Enabled  = false;
			this.cmbNEParroquia.Enabled = false;
			this.cmbNEBarrio.Enabled = false;
			this.txtCalle1.Enabled = false;
			this.txtNumcasa.Enabled = false;
			this.txtCalle2.Enabled = false;
			this.txtNETelefono.Enabled = false;
			this.txtNotas.Enabled  = false;
			this.txtNotas2.Enabled = false;
			this.cmdCalificacion.Enabled  = false;
			this.cmbEstado.Enabled = false;
			this.txtScore.Enabled = false;
			this.txtDocumento.Enabled = false;
		  this.btnGuardar.Enabled = false;

			#endregion Desabilitacontroles 
		}
		#region Focus 

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
	
		}

		private void txtDocumento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmdCalificacion.Focus();
		}

		private void cmdCalificacion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtScore.Focus();
		}

		

		private void cmbEstado_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtApellidoP.Focus();
		}

		private void txtApellidoP_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtApellidoM.Focus();
		}

		private void txtApellidoM_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtPnombre.Focus();
		}

		private void txtPnombre_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtSnombre.Focus();
		}

		private void txtSnombre_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbNEProvincia.Focus();
		}

		private void cmbNEProvincia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbNECanton.Focus();
		}

		private void cmbNECanton_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbNEParroquia.Focus();
		}

		private void cmbNEParroquia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbNEBarrio.Focus();
		}

		private void cmbNEBarrio_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtCalle1.Focus();
		}

		private void txtCalle1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtNumcasa.Focus();
		}

		private void txtNumcasa_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtCalle2.Focus();
		}

		private void txtCalle2_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtNETelefono.Focus();
		}

		private void txtNETelefono_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtNotas.Focus();
		}

		private void txtNotas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtNotas2.Focus();
		}

		private void txtNotas2_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}
#endregion Focus 

		private void cmbNEProvincia_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbNEProvincia.ActiveRow != null)
			{
				this.cmbNECanton.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 1, {0}, 0", (int) this.cmbNEProvincia.Value));
				
			}
		}

		private void cmbNECanton_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbNECanton.ActiveRow != null)
			{
				this.cmbNEParroquia.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 2, {0}, 0", (int) this.cmbNECanton.Value));
				
			}
		}

		private void cmbNEParroquia_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbNEParroquia.ActiveRow != null)
			{
				this.cmbNEBarrio.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 3, {0}, 0", (int) this.cmbNEParroquia.Value));
			}
		}

		private void txtDocumento_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.cmbTipoRuc.ActiveRow != null)
			{
				int iTipoRuc = 0;
				if ((int) this.cmbTipoRuc.Value == 1) iTipoRuc = 5;// cedula
				if ((int) this.cmbTipoRuc.Value == 2) iTipoRuc = 6;// pasaporte
				if ((int) this.cmbTipoRuc.Value == 3) iTipoRuc = 4;// ruc
				if (!Validacion.vbIdentificacion(this.txtDocumento, iTipoRuc, cdsSeteoF)) e.Cancel = true;
			}
		}

		private void txtDocumento_Validated(object sender, System.EventArgs e)
		{
			 bCarga = true;
			if (sCreSolicitud == this.txtDocumento.Text)
			{
				MessageBox.Show(string.Format("No puede ser el Cliente Garante de si mismo"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtDocumento.Focus();
				return;
			}
			
			if (idCre_Solicitud > 0 && bnuevo)
			{
				string ssqlDo= string.Format ("Exec Cre_ConsultaGarantes {0}", idCre_Solicitud ); 
				SqlDataReader drDO = Funcion.rEscalarSQL(cdsSeteoF,ssqlDo,true);
				drDO.Read();
				if(drDO.HasRows)
				{

					this.txtidGarante.Value = drDO.GetInt32 (0);
					this.txtIdCre_Solicitud.Value = drDO.GetInt32 (1);
					this.txtNotas2.Text = drDO.GetString (2);
				}
				drDO.Close();
			
			Consulta((int)this.txtidGarante.Value, (int) this.txtIdCre_Solicitud.Value);

			}
				 bCarga = false;
			
		}

		private void Consulta (int idGarante, int iCreSolicitud)
		{
			int iEstado = 0;

			int iCre_DatosGenerales = 0;
			

			if(idGarante > 0)
			{
				
				#region Carga Solicitudes

				string ssql= string.Format ("Exec Cre_ConsultaSolicitudes {0}",iCreSolicitud ); 
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF,ssql,true);
				dr.Read();
				if (dr.HasRows)
				{		
					this.txtIdCre_Solicitud.Value = dr.GetInt32 (0);

					
		
					#region Bloqueo Estado 2
					
				}
				#endregion Bloqueo Estado 2
				dr.Close();

				#endregion Carga Solicitudes
	
				#region Datos Generales

				string ssqDG= string.Format ("Exec Cre_ConsultaDatosGenerales {0},2",idGarante ); 
				SqlDataReader drDG = Funcion.rEscalarSQL(cdsSeteoF,ssqDG,true);
				drDG.Read();
				if (drDG.HasRows)
				{		
					this.txtIdCre_DatosGenerales.Value = drDG.GetInt32 (0);
					if (drDG.GetValue (3) != System.DBNull.Value)this.cmbTipoRuc.Value = drDG.GetInt32 (3);
					if (drDG.GetValue (4) != System.DBNull.Value)this.txtDocumento.Text = drDG.GetString (4);
					if (drDG.GetValue (5)!= System.DBNull.Value)this.txtApellidoP.Text = drDG.GetString (5);
					if (drDG.GetValue (6)!= System.DBNull.Value)this.txtApellidoM.Text = drDG.GetString (6);
					if (drDG.GetValue (7)!= System.DBNull.Value)this.txtPnombre.Text = drDG.GetString (7);
					if (drDG.GetValue (8) != System.DBNull.Value)this.txtSnombre.Text = drDG.GetString (8);
					if (drDG.GetValue (9) != System.DBNull.Value)this.cmdCalificacion.Value = drDG.GetInt32 (9);
					if (drDG.GetValue (13) != System.DBNull.Value)this.cmbEstado.Value = drDG.GetInt32 (13);
					if (drDG.GetValue (2) != System.DBNull.Value)this.txtIdCre_DatosGenerales.Value = drDG.GetInt32 (0);
					if (drDG.GetValue (11) != System.DBNull.Value)this.txtScore.Value = drDG.GetInt32 (11);

					//													this.cmdEstadoSolicitud.Value
				}
				drDG.Close();

				//				string ssqCDG= string.Format ("Exec Cre_ConsultaSolicitudes {0}",(int)this.txtidIdentificacion.Value ); 
				//				SqlDataReader drCDG = Funcion.rEscalarSQL(cdsSeteoF,ssqCDG,true);
				//				drCDG.Read();
				//				if (drCDG.HasRows)
				//				{
				//					if (drCDG.GetValue (9) != System.DBNull.Value)this.cmdEstadoSolicitud.Value = drCDG.GetInt32 (9);
				//				}
				//				drCDG.Close();

				//

				string sqlCDU= string.Format ("Exec Cre_ConsultaDatosUbicacion {0},2",idGarante ); 
				SqlDataReader drCub = Funcion.rEscalarSQL(cdsSeteoF,sqlCDU,true);
				drCub.Read();
				if (drCub.HasRows)
				{
					this.txtIdCre_DatosUbicacion.Value = drCub.GetInt32 (0);
					if (drCub.GetValue (3) != System.DBNull.Value)	this.cmbNEProvincia.Value = drCub.GetInt32 (3);
					if (drCub.GetValue (4) != System.DBNull.Value)this.cmbNECanton.Value = drCub.GetInt32 (4);
					if (drCub.GetValue (5)!= System.DBNull.Value)this.cmbNEParroquia.Value = drCub.GetInt32 (5);
					if (drCub.GetValue (6) != System.DBNull.Value)this.cmbNEBarrio.Value = drCub.GetInt32 (6);
					if (drCub.GetValue (7)!= System.DBNull.Value)this.txtCalle1.Text =  drCub.GetString (7);
					if (drCub.GetValue (8) != System.DBNull.Value)this.txtNumcasa.Text =  drCub.GetString (8);
					if (drCub.GetValue (8)!= System.DBNull.Value)this.txtCalle2.Text =  drCub.GetString (9);
					if (drCub.GetValue (10) != System.DBNull.Value)this.txtNotas.Text = drCub.GetString (10);
					if (drCub.GetValue (11)!= System.DBNull.Value)this.txtNETelefono.Text = drCub.GetString (11);
					if (drCub.GetValue (2) != System.DBNull.Value)this.txtIdCre_DatosUbicacion.Value = drCub.GetInt32 (0);
				}

				drCub.Close(); 
				//
			

				#endregion Datos Generales      

			}
		}

		private void txtCuotaAsignada_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
		
		}

		private void txtScore_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
						if(e.KeyCode == Keys.Enter) this.cmbEstado.Focus();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void cmbTipoRuc_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
		
		}

		private void txtDocumento_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region valida Campos Generales 

			if (!Validacion.vbComboVacio(this.cmbTipoRuc, "Seleccione el tipo de Documento ")) return;
			if (!Validacion.vbTexto(this.txtApellidoP, 5, 60, "Ingrese el Apellido Paterno")) return;
//			if (!Validacion.vbTexto(this.txtApellidoM, 5, 60, "Ingrese el Apellido Materno")) return;
			if (!Validacion.vbTexto(this.txtPnombre, 5, 60, "Ingrese el Primer Nombre")) return;
//			if (!Validacion.vbTexto(this.txtSnombre, 5, 60, "Ingrese el Segundo nombre")) return;
			if (!Validacion.vbFechaEnDocumentos(this.dtpNEFechaNacimiento, "Ingrese la fecha de Nacimiento", true, "Cobros", cdsSeteoF)) return;
			if (!Validacion.vbComboVacio(this.cmbNEProvincia, "Seleccione la Provincia ")) return;
			if (!Validacion.vbComboVacio(this.cmbNECanton, "Seleccione el Canton ")) return;
			if (!Validacion.vbComboVacio(this.cmbNEParroquia, "Seleccione la Parroquia")) return;
			if (!Validacion.vbComboVacio(this.cmbNEBarrio, "Seleccione el Barrio ")) return;
			if (!Validacion.vbTexto(this.txtCalle1, 5, 70, "Ingrese la calle Principal")) return;
			if (!Validacion.vbTexto(this.txtNumcasa, 5, 30, "Ingrese el número de casa")) return;
			if (!Validacion.vbTexto(this.txtCalle2, 5, 70, "Ingrese la calle Secundaria")) return;

			#endregion valida Campos Generales 

			#region Valida Telefono
			if (this.txtNETelefono.Text.ToString().Trim().Length == 0   )
			{
				MessageBox.Show(string.Format("Debe Ingresar el Telefono Convencional o el Telefono Celular"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtNETelefono.Focus();
		
				return;	
			}
			if (!Validacion.vbValidaTelefono(this.txtNETelefono, 1, false)) return;

			#endregion Valida Telefono

			#region Valida cAMPO

			if (!Validacion.vbTexto(this.txtNotas, 5, 100, "Ingrese la Referencia de Ubicación")) return;
			if (!Validacion.vbTexto(this.txtNotas2, 5, 100, "Ingrese una Observación")) return;
			if (!Validacion.vbComboVacio(this.cmdCalificacion, "Seleccione la Calificación ")) return;
			if (!Validacion.vbComboVacio(this.cmbEstado, "Seleccione un Estado ")) return;
    	if (!Validacion.vbCampoEnteroVacio(this.txtScore, "Ingrese el Score",1,999)) return;


			#endregion Valida cAMPO

			#region Valida Cuota 


//			if ((int)this.txtCuotaAsignada.Value <= 100)
//			{
//				MessageBox.Show(" La Cuota asignada no puede ser menor a 100", "Point Tehnology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//				this.txtCuotaAsignada.Focus();
//				return;
//			}
     
//			if ((int)this.cmdEntidad.Value == 2)
//			{
//				if ((int)this.txtCuotaAsignada.Value > 2100)
//				{
//					MessageBox.Show(" La Cuota asignada maxima para Credito Solidario es 2100", "Point Tehnology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//					this.txtCuotaAsignada.Focus();
//					return;
//
//				}
//			}
//			if ((int)this.cmdEntidad.Value == 4)
//			{
//				if ((int)this.txtCuotaAsignada.Value > 2000)
//				{
//					MessageBox.Show(" La Cuota asignada maxima para Credito Pichincha es 2000", "Point Tehnology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//					this.txtCuotaAsignada.Focus();
//					return;
//
//				}
//			}
			#endregion Valida Cuota 

			
			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + ";Max Pool Size=75000;"))
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
			
			   #region Guarda Cuota

					string sSQLDetalle = string.Format("Exec Cre_GuardaGarantes {0}, {1}, '{2}', '{3}'",
						(int)this.txtidGarante.Value, 
						idCre_Solicitud, 
						Convert.ToDateTime(this.dtpNEFechaNacimiento.Value).ToString("yyyyMMdd"),
						this.txtNotas2.Text);
					oCmdActualiza.CommandText = sSQLDetalle;
					this.txtidGarante.Value = (int)oCmdActualiza.ExecuteScalar();

					#endregion Guarda Cuota

					string sSQLDGenerales = string.Format("Exec Cre_GuardadDGeneralesRevisionCedulas  {0}, {1}, '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', {8}, {9}, {10},{11}, {12}",
						
						(int)this.txtIdCre_DatosGenerales.Value, 
						(int)this.cmbTipoRuc.Value, 
						this.txtDocumento.Text,
						this.txtApellidoP.Text , 
						this.txtApellidoM.Text, 
						this.txtPnombre.Text, 
						this.txtSnombre.Text,
						Convert.ToDateTime(this.dtpNEFechaNacimiento.Value).ToString("yyyyMMdd"), 
						(int)this.cmdCalificacion.Value, 
						(int)this.cmbEstado.Value,
						(int)this.txtidGarante.Value, 2, (int)this.txtScore.Value);

					oCmdActualiza.CommandText = sSQLDGenerales;
					this.txtIdCre_DatosGenerales.Value = (int)oCmdActualiza.ExecuteScalar();

					string sSQLDUbicacion = string.Format("Exec GuardaDatosUbicacion  {0}, {1}, {2}, {3}, {4}, '{5}', '{6}', '{7}', '{8}', '{9}', '{10}', {11},{12}",
						
						this.txtIdCre_DatosUbicacion.Value,
						this.cmbNEProvincia.Value,
						this.cmbNECanton.Value, 
						this.cmbNEParroquia.Value, 
						this.cmbNEBarrio.Value ,
						this.txtCalle1.Text ,
						this.txtNumcasa.Text,
						this.txtCalle2.Text ,
						this.txtNotas.Text ,
						this.txtNETelefono.Text,
						this.txtNETelefono.Text ,
						this.txtidGarante.Value, 2 );

					oCmdActualiza.CommandText = sSQLDUbicacion;
					
					this.txtIdCre_DatosUbicacion.Value = (int)oCmdActualiza.ExecuteScalar();

					oTransaction.Commit();

				}
				catch (Exception ex)
				{
					oTransaction.Rollback();

					MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);					
				}
				finally
				{
					oConexion.Close();
				}
			}
			
     this.DialogResult = DialogResult.OK;

		}

		private void txtNETelefono_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}
	
		class validar 
		{
			public static void Solonumeros (KeyPressEventArgs Nu)
			{
				if(char.IsDigit(Nu.KeyChar))
				{
			
					Nu.Handled = false;
			
				}
				else if (char.IsControl(Nu.KeyChar))
				{
			
					Nu.Handled = false;			} 
				else
				{
					Nu.Handled= true;
				}
			}
		}
	}
}
