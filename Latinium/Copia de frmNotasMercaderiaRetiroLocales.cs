using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmNotasMercaderiaRetiroLocales.
	/// </summary>
	public class frmNotasMercaderiaRetiroLocales : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtpFecha;
		private System.Windows.Forms.Label lblNota;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNota;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Label lblTransportista;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label lblEstado;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtBodegas;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbTransportista;
		private System.Windows.Forms.Label lblPersonal;
		private System.Windows.Forms.GroupBox groupBox1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmNotasMercaderiaRetiroLocales()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmNotasMercaderiaRetiroLocales));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			this.label1 = new System.Windows.Forms.Label();
			this.dtpFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblNota = new System.Windows.Forms.Label();
			this.txtNota = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.lblTransportista = new System.Windows.Forms.Label();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.lblEstado = new System.Windows.Forms.Label();
			this.txtBodegas = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbTransportista = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblPersonal = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.dtpFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNota)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTransportista)).BeginInit();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(8, 10);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(36, 16);
			this.label1.TabIndex = 658;
			this.label1.Text = "Fecha";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtpFecha
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtpFecha.Appearance = appearance1;
			this.dtpFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtpFecha.DateButtons.Add(dateButton1);
			this.dtpFecha.Enabled = false;
			this.dtpFecha.Format = "dd/MM/yyyy";
			this.dtpFecha.Location = new System.Drawing.Point(96, 8);
			this.dtpFecha.Name = "dtpFecha";
			this.dtpFecha.NonAutoSizeHeight = 23;
			this.dtpFecha.Size = new System.Drawing.Size(112, 21);
			this.dtpFecha.SpinButtonsVisible = true;
			this.dtpFecha.TabIndex = 657;
			this.dtpFecha.Value = ((object)(resources.GetObject("dtpFecha.Value")));
			this.dtpFecha.ValueChanged += new System.EventHandler(this.dtpFecha_ValueChanged);
			// 
			// lblNota
			// 
			this.lblNota.AutoSize = true;
			this.lblNota.BackColor = System.Drawing.Color.Transparent;
			this.lblNota.Location = new System.Drawing.Point(-60, 125);
			this.lblNota.Name = "lblNota";
			this.lblNota.Size = new System.Drawing.Size(34, 16);
			this.lblNota.TabIndex = 704;
			this.lblNota.Text = "Notas";
			this.lblNota.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNota
			// 
			this.txtNota.AcceptsReturn = true;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNota.Appearance = appearance2;
			this.txtNota.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNota.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNota.Location = new System.Drawing.Point(96, 72);
			this.txtNota.MaxLength = 300;
			this.txtNota.Multiline = true;
			this.txtNota.Name = "txtNota";
			this.txtNota.Size = new System.Drawing.Size(300, 66);
			this.txtNota.TabIndex = 703;			
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(8, 97);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(28, 16);
			this.label2.TabIndex = 705;
			this.label2.Text = "Nota";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(8, 160);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 745;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(96, 160);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 746;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// lblTransportista
			// 
			this.lblTransportista.AutoSize = true;
			this.lblTransportista.BackColor = System.Drawing.Color.Transparent;
			this.lblTransportista.Location = new System.Drawing.Point(8, 43);
			this.lblTransportista.Name = "lblTransportista";
			this.lblTransportista.Size = new System.Drawing.Size(70, 16);
			this.lblTransportista.TabIndex = 756;
			this.lblTransportista.Text = "Transportista";
			this.lblTransportista.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblTransportista.Visible = false;
			// 
			// txtEstado
			// 
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEstado.Appearance = appearance3;
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.FormatString = "###0";
			this.txtEstado.Location = new System.Drawing.Point(344, 8);
			this.txtEstado.MaxValue = 100000000;
			this.txtEstado.MinValue = 0;
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(24, 21);
			this.txtEstado.TabIndex = 757;
			this.txtEstado.Visible = false;
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
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Tahoma", 9F);
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.lblEstado.Location = new System.Drawing.Point(296, 66);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 18);
			this.lblEstado.TabIndex = 760;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtBodegas
			// 
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBodegas.Appearance = appearance4;
			this.txtBodegas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBodegas.Enabled = false;
			this.txtBodegas.FormatString = "###0";
			this.txtBodegas.Location = new System.Drawing.Point(376, 8);
			this.txtBodegas.MaxValue = 100000000;
			this.txtBodegas.MinValue = 0;
			this.txtBodegas.Name = "txtBodegas";
			this.txtBodegas.PromptChar = ' ';
			this.txtBodegas.Size = new System.Drawing.Size(24, 21);
			this.txtBodegas.TabIndex = 761;
			this.txtBodegas.Visible = false;
			// 
			// cmbTransportista
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTransportista.Appearance = appearance5;
			this.cmbTransportista.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTransportista.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 335;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbTransportista.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbTransportista.DisplayMember = "Nombre";
			this.cmbTransportista.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTransportista.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTransportista.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTransportista.Location = new System.Drawing.Point(96, 40);
			this.cmbTransportista.Name = "cmbTransportista";
			this.cmbTransportista.Size = new System.Drawing.Size(240, 22);
			this.cmbTransportista.TabIndex = 762;
			this.cmbTransportista.ValueMember = "idPersonal";
			// 
			// lblPersonal
			// 
			this.lblPersonal.AutoSize = true;
			this.lblPersonal.BackColor = System.Drawing.Color.Transparent;
			this.lblPersonal.Location = new System.Drawing.Point(8, 43);
			this.lblPersonal.Name = "lblPersonal";
			this.lblPersonal.Size = new System.Drawing.Size(49, 16);
			this.lblPersonal.TabIndex = 763;
			this.lblPersonal.Text = "Personal";
			this.lblPersonal.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblPersonal.Visible = false;			
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 144);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(416, 8);
			this.groupBox1.TabIndex = 764;
			this.groupBox1.TabStop = false;
			// 
			// frmNotasMercaderiaRetiroLocales
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(410, 192);
			this.ControlBox = false;
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.lblPersonal);
			this.Controls.Add(this.lblTransportista);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.lblNota);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbTransportista);
			this.Controls.Add(this.txtBodegas);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.txtNota);
			this.Controls.Add(this.dtpFecha);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.Name = "frmNotasMercaderiaRetiroLocales";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Notas";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmNotasMercaderiaRetiroLocales_KeyDown);
			this.Load += new System.EventHandler(this.frmNotasMercaderiaRetiroLocales_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtpFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNota)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTransportista)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		
		private void dtpFecha_ValueChanged(object sender, System.EventArgs e)
		{
			//			this.cmbTransportista.DataSource = FuncionesProcedimientos.Lista_Vendedores(Convert.ToDateTime(this.dtpFecha.Value), (int)this.txtBodegas.Value, cdsSeteoF);
			//	      this.cmbRecibidopor.DataSource = FuncionesProcedimientos.Lista_Vendedores(Convert.ToDateTime(this.dtFecha.Value), (int)this.cmbBodega.Value, cdsSeteoF);
		}

		private void frmNotasMercaderiaRetiroLocales_Load(object sender, System.EventArgs e)
		{
			this.dtpFecha.Value = DateTime.Today;
			this.cmbTransportista.DataSource = FuncionesProcedimientos.Lista_Vendedores(Convert.ToDateTime(this.dtpFecha.Value), (int)this.txtBodegas.Value, cdsSeteoF);
			
			if( (int) this.txtEstado.Value == 1)
			{
				this.lblEstado.Text ="RECIBIDO";
				this.lblTransportista.Visible = true;

			}
			if( (int) this.txtEstado.Value == 2)
			{
				this.lblPersonal.Visible = true;
				this.lblEstado.Text ="ENVIADO A MATRIZ";
			}
			if( (int) this.txtEstado.Value == 3)
			{
				this.lblEstado.Text ="RECIBIDO EN MATRIZ ";
			}
			if( (int) this.txtEstado.Value == 4)
			{
				this.cmbTransportista.Visible= false;
				this.lblTransportista.Visible = false;
				this.lblEstado.Text ="ANULANDO";
			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if( (int) this.txtEstado.Value == 1)
			{
				this.lblEstado.Text ="ENVIADO A MATRIZ";
				if (!Validacion.vbComboVacio(this.cmbTransportista, "Seleccione un Transportista")) return;
			}
			if( (int) this.txtEstado.Value == 2)
			{
				this.lblEstado.Text ="ENVIADO A MATRIZ";
				if (!Validacion.vbComboVacio(this.cmbTransportista, "Seleccione el Personal que recibe el producto ")) return;
			}
			if( (int) this.txtEstado.Value == 3)
			{
				this.lblEstado.Text ="RECIBIDO EN MATRIZ ";
			}

			if( (int) this.txtEstado.Value == 4)
			{
				this.cmbTransportista.Visible= false;
				this.lblTransportista.Visible = false;
				this.lblEstado.Text ="ANULADO";
			}

			if (!Validacion.vbTexto(this.txtNota, 20, 1000, "Notas")) return;
	     
			this.DialogResult = DialogResult.OK;
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void frmNotasMercaderiaRetiroLocales_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}
	}
}
