using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmcatalogo.
	/// </summary>
	public class frmcatalogo : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidArticulo;
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.Label lblCliente;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinEditors.UltraComboEditor cmbTipo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdPromocion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcion;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFin;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtInicio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidHoraPromocion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor2;
		private System.Windows.Forms.Label lblTipoCatalogo;
		private System.Windows.Forms.Label label2;
		private DevExpress.XtraEditors.SimpleButton btnActualizaPrecios;
		private DevExpress.XtraEditors.SimpleButton btFinalizar;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private DevExpress.XtraEditors.SimpleButton btBuscar;
		private DevExpress.XtraEditors.SimpleButton btGrabar;
		private DevExpress.XtraEditors.SimpleButton btNuevo;
		private DevExpress.XtraEditors.SimpleButton btSalir;
		private DevExpress.XtraEditors.SimpleButton btEditar;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.TextBox txtpaginacatalogo;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmcatalogo()
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
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmcatalogo));
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.txtidArticulo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label16 = new System.Windows.Forms.Label();
			this.lblCliente = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.cmbTipo = new Infragistics.Win.UltraWinEditors.UltraComboEditor();
			this.txtIdPromocion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDescripcion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtFin = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtInicio = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtidHoraPromocion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblTipoCatalogo = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.btnActualizaPrecios = new DevExpress.XtraEditors.SimpleButton();
			this.btFinalizar = new DevExpress.XtraEditors.SimpleButton();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.btBuscar = new DevExpress.XtraEditors.SimpleButton();
			this.btGrabar = new DevExpress.XtraEditors.SimpleButton();
			this.btNuevo = new DevExpress.XtraEditors.SimpleButton();
			this.btSalir = new DevExpress.XtraEditors.SimpleButton();
			this.btEditar = new DevExpress.XtraEditors.SimpleButton();
			this.label3 = new System.Windows.Forms.Label();
			this.txtpaginacatalogo = new System.Windows.Forms.TextBox();
			((System.ComponentModel.ISupportInitialize)(this.txtidArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdPromocion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFin)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtInicio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidHoraPromocion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor2)).BeginInit();
			this.SuspendLayout();
			// 
			// txtidArticulo
			// 
			this.txtidArticulo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidArticulo.Location = new System.Drawing.Point(1184, 16);
			this.txtidArticulo.Name = "txtidArticulo";
			this.txtidArticulo.PromptChar = ' ';
			this.txtidArticulo.Size = new System.Drawing.Size(16, 21);
			this.txtidArticulo.TabIndex = 309;
			this.txtidArticulo.Visible = false;
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.Enabled = false;
			this.label16.Location = new System.Drawing.Point(-152, 16);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(26, 16);
			this.label16.TabIndex = 308;
			this.label16.Text = "Tipo";
			this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Enabled = false;
			this.lblCliente.Location = new System.Drawing.Point(8, 48);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(93, 16);
			this.lblCliente.TabIndex = 305;
			this.lblCliente.Text = "Nombre Catálogo";
			this.lblCliente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Enabled = false;
			this.label1.Location = new System.Drawing.Point(616, 16);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(37, 16);
			this.label1.TabIndex = 304;
			this.label1.Text = "Hasta:";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Enabled = false;
			this.label6.Location = new System.Drawing.Point(432, 16);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(40, 16);
			this.label6.TabIndex = 303;
			this.label6.Text = "Desde:";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbTipo
			// 
			appearance1.FontData.Name = "Tahoma";
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipo.Appearance = appearance1;
			this.cmbTipo.DropDownStyle = Infragistics.Win.DropDownStyle.DropDownList;
			this.cmbTipo.Enabled = false;
			valueListItem1.DataValue = "ValueListItem0";
			valueListItem1.DisplayText = "PROMOCIÓN";
			this.cmbTipo.Items.Add(valueListItem1);
			this.cmbTipo.Location = new System.Drawing.Point(120, 16);
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(160, 22);
			this.cmbTipo.TabIndex = 307;
			// 
			// txtIdPromocion
			// 
			this.txtIdPromocion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdPromocion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdPromocion.Location = new System.Drawing.Point(1152, 16);
			this.txtIdPromocion.Name = "txtIdPromocion";
			this.txtIdPromocion.PromptChar = ' ';
			this.txtIdPromocion.Size = new System.Drawing.Size(24, 21);
			this.txtIdPromocion.TabIndex = 306;
			this.txtIdPromocion.Visible = false;
			// 
			// txtDescripcion
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescripcion.Appearance = appearance2;
			this.txtDescripcion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDescripcion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescripcion.Enabled = false;
			this.txtDescripcion.Location = new System.Drawing.Point(120, 48);
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.Size = new System.Drawing.Size(240, 21);
			this.txtDescripcion.TabIndex = 302;
			// 
			// dtFin
			// 
			appearance3.FontData.Name = "Tahoma";
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFin.Appearance = appearance3;
			this.dtFin.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFin.DateButtons.Add(dateButton1);
			this.dtFin.Enabled = false;
			this.dtFin.Format = "dd/MM/yyyy HH:mm";
			this.dtFin.Location = new System.Drawing.Point(656, 16);
			this.dtFin.Name = "dtFin";
			this.dtFin.NonAutoSizeHeight = 23;
			this.dtFin.Size = new System.Drawing.Size(144, 21);
			this.dtFin.SpinButtonsVisible = true;
			this.dtFin.TabIndex = 301;
			this.dtFin.Value = ((object)(resources.GetObject("dtFin.Value")));
			// 
			// dtInicio
			// 
			appearance4.FontData.Name = "Tahoma";
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtInicio.Appearance = appearance4;
			this.dtInicio.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtInicio.DateButtons.Add(dateButton2);
			this.dtInicio.Enabled = false;
			this.dtInicio.Format = "dd/MM/yyyy HH:mm";
			this.dtInicio.Location = new System.Drawing.Point(472, 16);
			this.dtInicio.Name = "dtInicio";
			this.dtInicio.NonAutoSizeHeight = 23;
			this.dtInicio.Size = new System.Drawing.Size(144, 21);
			this.dtInicio.SpinButtonsVisible = true;
			this.dtInicio.TabIndex = 300;
			this.dtInicio.Value = ((object)(resources.GetObject("dtInicio.Value")));
			// 
			// txtidHoraPromocion
			// 
			this.txtidHoraPromocion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidHoraPromocion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidHoraPromocion.Enabled = false;
			this.txtidHoraPromocion.Location = new System.Drawing.Point(928, 16);
			this.txtidHoraPromocion.Name = "txtidHoraPromocion";
			this.txtidHoraPromocion.PromptChar = ' ';
			this.txtidHoraPromocion.Size = new System.Drawing.Size(8, 21);
			this.txtidHoraPromocion.TabIndex = 335;
			// 
			// ultraNumericEditor1
			// 
			this.ultraNumericEditor1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.ultraNumericEditor1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor1.Location = new System.Drawing.Point(936, 16);
			this.ultraNumericEditor1.Name = "ultraNumericEditor1";
			this.ultraNumericEditor1.PromptChar = ' ';
			this.ultraNumericEditor1.Size = new System.Drawing.Size(16, 21);
			this.ultraNumericEditor1.TabIndex = 334;
			this.ultraNumericEditor1.Visible = false;
			// 
			// ultraNumericEditor2
			// 
			this.ultraNumericEditor2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.ultraNumericEditor2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor2.Location = new System.Drawing.Point(952, 16);
			this.ultraNumericEditor2.Name = "ultraNumericEditor2";
			this.ultraNumericEditor2.PromptChar = ' ';
			this.ultraNumericEditor2.Size = new System.Drawing.Size(24, 21);
			this.ultraNumericEditor2.TabIndex = 333;
			this.ultraNumericEditor2.Visible = false;
			// 
			// lblTipoCatalogo
			// 
			this.lblTipoCatalogo.Location = new System.Drawing.Point(8, 16);
			this.lblTipoCatalogo.Name = "lblTipoCatalogo";
			this.lblTipoCatalogo.Size = new System.Drawing.Size(96, 23);
			this.lblTipoCatalogo.TabIndex = 343;
			this.lblTipoCatalogo.Text = "Tipo de Cátalogo:";
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(296, 16);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(136, 23);
			this.label2.TabIndex = 344;
			this.label2.Text = "Vigencia del Catálogo:";
			// 
			// btnActualizaPrecios
			// 
			this.btnActualizaPrecios.Appearance.Options.UseTextOptions = true;
			this.btnActualizaPrecios.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
			this.btnActualizaPrecios.Enabled = false;
			this.btnActualizaPrecios.Image = ((System.Drawing.Image)(resources.GetObject("btnActualizaPrecios.Image")));
			this.btnActualizaPrecios.ImageAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.btnActualizaPrecios.Location = new System.Drawing.Point(952, 248);
			this.btnActualizaPrecios.Name = "btnActualizaPrecios";
			this.btnActualizaPrecios.Size = new System.Drawing.Size(96, 24);
			this.btnActualizaPrecios.TabIndex = 352;
			this.btnActualizaPrecios.Text = "&Precios";
			this.btnActualizaPrecios.Visible = false;
			// 
			// btFinalizar
			// 
			this.btFinalizar.Appearance.Options.UseTextOptions = true;
			this.btFinalizar.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
			this.btFinalizar.Enabled = false;
			this.btFinalizar.Image = ((System.Drawing.Image)(resources.GetObject("btFinalizar.Image")));
			this.btFinalizar.ImageAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.btFinalizar.Location = new System.Drawing.Point(952, 176);
			this.btFinalizar.Name = "btFinalizar";
			this.btFinalizar.Size = new System.Drawing.Size(96, 24);
			this.btFinalizar.TabIndex = 351;
			this.btFinalizar.Text = "&Finalizar";
			// 
			// btCancelar
			// 
			this.btCancelar.Appearance.Options.UseTextOptions = true;
			this.btCancelar.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
			this.btCancelar.Enabled = false;
			this.btCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btCancelar.Image")));
			this.btCancelar.ImageAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.btCancelar.ImageIndex = 61;
			this.btCancelar.Location = new System.Drawing.Point(952, 200);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(96, 24);
			this.btCancelar.TabIndex = 350;
			this.btCancelar.Text = "&Cancelar";
			// 
			// btBuscar
			// 
			this.btBuscar.Appearance.Options.UseTextOptions = true;
			this.btBuscar.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
			this.btBuscar.Image = ((System.Drawing.Image)(resources.GetObject("btBuscar.Image")));
			this.btBuscar.ImageAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.btBuscar.ImageIndex = 57;
			this.btBuscar.Location = new System.Drawing.Point(952, 104);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(96, 24);
			this.btBuscar.TabIndex = 349;
			this.btBuscar.Text = "&Buscar";
			// 
			// btGrabar
			// 
			this.btGrabar.Appearance.Options.UseTextOptions = true;
			this.btGrabar.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
			this.btGrabar.Enabled = false;
			this.btGrabar.Image = ((System.Drawing.Image)(resources.GetObject("btGrabar.Image")));
			this.btGrabar.ImageAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.btGrabar.ImageIndex = 9;
			this.btGrabar.Location = new System.Drawing.Point(952, 152);
			this.btGrabar.Name = "btGrabar";
			this.btGrabar.Size = new System.Drawing.Size(96, 24);
			this.btGrabar.TabIndex = 348;
			this.btGrabar.Text = "&Grabar";
			// 
			// btNuevo
			// 
			this.btNuevo.Appearance.Options.UseTextOptions = true;
			this.btNuevo.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
			this.btNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btNuevo.Image")));
			this.btNuevo.ImageAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.btNuevo.ImageIndex = 8;
			this.btNuevo.Location = new System.Drawing.Point(952, 80);
			this.btNuevo.Name = "btNuevo";
			this.btNuevo.Size = new System.Drawing.Size(96, 24);
			this.btNuevo.TabIndex = 347;
			this.btNuevo.Text = "&Nuevo";
			// 
			// btSalir
			// 
			this.btSalir.Appearance.Options.UseTextOptions = true;
			this.btSalir.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
			this.btSalir.Image = ((System.Drawing.Image)(resources.GetObject("btSalir.Image")));
			this.btSalir.ImageAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.btSalir.ImageIndex = 62;
			this.btSalir.Location = new System.Drawing.Point(952, 224);
			this.btSalir.Name = "btSalir";
			this.btSalir.Size = new System.Drawing.Size(96, 24);
			this.btSalir.TabIndex = 346;
			this.btSalir.Text = "&Salir";
			// 
			// btEditar
			// 
			this.btEditar.Appearance.Options.UseTextOptions = true;
			this.btEditar.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
			this.btEditar.Enabled = false;
			this.btEditar.Image = ((System.Drawing.Image)(resources.GetObject("btEditar.Image")));
			this.btEditar.ImageAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.btEditar.ImageIndex = 56;
			this.btEditar.Location = new System.Drawing.Point(952, 128);
			this.btEditar.Name = "btEditar";
			this.btEditar.Size = new System.Drawing.Size(96, 24);
			this.btEditar.TabIndex = 345;
			this.btEditar.Text = "&Editar";
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(8, 88);
			this.label3.Name = "label3";
			this.label3.TabIndex = 353;
			this.label3.Text = "Paginas Catálogo:";
			this.label3.Click += new System.EventHandler(this.label3_Click);
			// 
			// txtpaginacatalogo
			// 
			this.txtpaginacatalogo.Location = new System.Drawing.Point(120, 88);
			this.txtpaginacatalogo.Name = "txtpaginacatalogo";
			this.txtpaginacatalogo.Size = new System.Drawing.Size(112, 20);
			this.txtpaginacatalogo.TabIndex = 354;
			this.txtpaginacatalogo.Text = "";
			// 
			// frmcatalogo
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1056, 630);
			this.Controls.Add(this.txtpaginacatalogo);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.btnActualizaPrecios);
			this.Controls.Add(this.btFinalizar);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btBuscar);
			this.Controls.Add(this.btGrabar);
			this.Controls.Add(this.btNuevo);
			this.Controls.Add(this.btSalir);
			this.Controls.Add(this.btEditar);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.lblTipoCatalogo);
			this.Controls.Add(this.txtidHoraPromocion);
			this.Controls.Add(this.ultraNumericEditor1);
			this.Controls.Add(this.ultraNumericEditor2);
			this.Controls.Add(this.txtidArticulo);
			this.Controls.Add(this.label16);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.cmbTipo);
			this.Controls.Add(this.txtIdPromocion);
			this.Controls.Add(this.txtDescripcion);
			this.Controls.Add(this.dtFin);
			this.Controls.Add(this.dtInicio);
			this.Name = "frmcatalogo";
			this.Text = "frmcatalogo";
			this.Load += new System.EventHandler(this.frmcatalogo_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtidArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdPromocion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFin)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtInicio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidHoraPromocion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor2)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmcatalogo_Load(object sender, System.EventArgs e)
		{
		
		}

		private void label3_Click(object sender, System.EventArgs e)
		{
		
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
