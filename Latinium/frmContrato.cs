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
	/// Descripción breve de frmContrato.
	/// </summary>
	public class frmContrato : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label lblFecha;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Label lblCliente;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private System.Windows.Forms.Label lblDetalle;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidProforma;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCliente;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtArchivo;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidcontrato;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Label lbltitulo;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtTitulo;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		int idcliente =0;
		int iestado = 0;
		int iproforma = 0;

		private C1.Data.C1DataSet cdsSeteoF;
		DateTime dtfecha ;
		string snombre;
		int idContrato;

		public frmContrato(int IdCliente, int iEstados, int iProforma, int idcontrato, DateTime dtFecha, string sNombre)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			 idcliente = IdCliente;
			 iestado = iEstados;
			 iproforma = iProforma;
			 idContrato = idcontrato;
			 dtfecha = dtFecha;
		  	snombre = sNombre;
			idContrato = idcontrato;
			
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmContrato));
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			this.txtidProforma = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblFecha = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblCliente = new System.Windows.Forms.Label();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtIdCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblDetalle = new System.Windows.Forms.Label();
			this.txtArchivo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtidcontrato = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.lbltitulo = new System.Windows.Forms.Label();
			this.txtTitulo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.txtidProforma)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArchivo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidcontrato)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTitulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// txtidProforma
			// 
			appearance1.ForeColor = System.Drawing.Color.Black;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtidProforma.Appearance = appearance1;
			this.txtidProforma.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidProforma.Enabled = false;
			this.txtidProforma.FormatString = "###0";
			this.txtidProforma.Location = new System.Drawing.Point(416, 48);
			this.txtidProforma.MaxValue = 100000000;
			this.txtidProforma.MinValue = 0;
			this.txtidProforma.Name = "txtidProforma";
			this.txtidProforma.PromptChar = ' ';
			this.txtidProforma.Size = new System.Drawing.Size(48, 21);
			this.txtidProforma.TabIndex = 700;
			this.txtidProforma.Visible = false;
			// 
			// lblFecha
			// 
			this.lblFecha.AutoSize = true;
			this.lblFecha.BackColor = System.Drawing.Color.Transparent;
			this.lblFecha.Location = new System.Drawing.Point(16, 16);
			this.lblFecha.Name = "lblFecha";
			this.lblFecha.Size = new System.Drawing.Size(36, 16);
			this.lblFecha.TabIndex = 704;
			this.lblFecha.Text = "Fecha";
			this.lblFecha.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance2;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton1.Caption = "Today";
			dateButton1.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(72, 16);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(104, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 705;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(16, 56);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(56, 16);
			this.lblCliente.TabIndex = 784;
			this.lblCliente.Text = "Proveedor";
			this.lblCliente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNombre
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance3;
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(72, 54);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(248, 21);
			this.txtNombre.TabIndex = 783;
			// 
			// txtIdCliente
			// 
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIdCliente.Appearance = appearance4;
			this.txtIdCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCliente.Enabled = false;
			this.txtIdCliente.FormatString = "###0";
			this.txtIdCliente.Location = new System.Drawing.Point(416, 128);
			this.txtIdCliente.MaxValue = 100000000;
			this.txtIdCliente.MinValue = 0;
			this.txtIdCliente.Name = "txtIdCliente";
			this.txtIdCliente.PromptChar = ' ';
			this.txtIdCliente.Size = new System.Drawing.Size(48, 21);
			this.txtIdCliente.TabIndex = 785;
			this.txtIdCliente.Visible = false;
			// 
			// txtEstado
			// 
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEstado.Appearance = appearance5;
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.FormatString = "###0";
			this.txtEstado.Location = new System.Drawing.Point(416, 88);
			this.txtEstado.MaxValue = 100000000;
			this.txtEstado.MinValue = 0;
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(48, 21);
			this.txtEstado.TabIndex = 786;
			this.txtEstado.Visible = false;
			// 
			// lblDetalle
			// 
			this.lblDetalle.AutoSize = true;
			this.lblDetalle.BackColor = System.Drawing.Color.Transparent;
			this.lblDetalle.Location = new System.Drawing.Point(16, 160);
			this.lblDetalle.Name = "lblDetalle";
			this.lblDetalle.Size = new System.Drawing.Size(34, 16);
			this.lblDetalle.TabIndex = 788;
			this.lblDetalle.Text = "Notas";
			this.lblDetalle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtArchivo
			// 
			this.txtArchivo.AcceptsReturn = true;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtArchivo.Appearance = appearance6;
			this.txtArchivo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtArchivo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtArchivo.Enabled = false;
			this.txtArchivo.Location = new System.Drawing.Point(72, 144);
			this.txtArchivo.MaxLength = 1000;
			this.txtArchivo.Multiline = true;
			this.txtArchivo.Name = "txtArchivo";
			this.txtArchivo.Size = new System.Drawing.Size(336, 64);
			this.txtArchivo.TabIndex = 787;
			// 
			// txtidcontrato
			// 
			appearance7.ForeColor = System.Drawing.Color.Black;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtidcontrato.Appearance = appearance7;
			this.txtidcontrato.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidcontrato.Enabled = false;
			this.txtidcontrato.FormatString = "###0";
			this.txtidcontrato.Location = new System.Drawing.Point(416, 8);
			this.txtidcontrato.MaxValue = 100000000;
			this.txtidcontrato.MinValue = 0;
			this.txtidcontrato.Name = "txtidcontrato";
			this.txtidcontrato.PromptChar = ' ';
			this.txtidcontrato.Size = new System.Drawing.Size(48, 21);
			this.txtidcontrato.TabIndex = 789;
			this.txtidcontrato.Visible = false;
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Tahoma", 11F, System.Drawing.FontStyle.Bold);
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.lblEstado.Location = new System.Drawing.Point(200, 16);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 21);
			this.lblEstado.TabIndex = 790;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnGuardar
			// 
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(72, 224);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 791;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnCancelar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(160, 224);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 792;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// lbltitulo
			// 
			this.lbltitulo.AutoSize = true;
			this.lbltitulo.Location = new System.Drawing.Point(16, 96);
			this.lbltitulo.Name = "lbltitulo";
			this.lbltitulo.Size = new System.Drawing.Size(32, 16);
			this.lbltitulo.TabIndex = 794;
			this.lbltitulo.Text = "Titulo";
			this.lbltitulo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTitulo
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTitulo.Appearance = appearance8;
			this.txtTitulo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtTitulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTitulo.Enabled = false;
			this.txtTitulo.Location = new System.Drawing.Point(72, 96);
			this.txtTitulo.Name = "txtTitulo";
			this.txtTitulo.Size = new System.Drawing.Size(336, 21);
			this.txtTitulo.TabIndex = 793;
			this.txtTitulo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtTitulo_KeyDown);
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
			// frmContrato
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(480, 270);
			this.Controls.Add(this.lbltitulo);
			this.Controls.Add(this.lblDetalle);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.lblFecha);
			this.Controls.Add(this.txtTitulo);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtidcontrato);
			this.Controls.Add(this.txtArchivo);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtIdCliente);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.txtidProforma);
			this.Name = "frmContrato";
			this.Text = "frmContrato";
			this.Load += new System.EventHandler(this.frmContrato_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtidProforma)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArchivo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidcontrato)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTitulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmContrato_Load(object sender, System.EventArgs e)
		{
			if (iproforma > 0)
			{
				#region Consulta Individual
				string ssql= string.Format ("Exec ConsultaIndividualContrato {0}", iproforma); 
				SqlDataReader drr = Funcion.rEscalarSQL(cdsSeteoF,ssql,true);
				drr.Read();

        this.txtidcontrato.Value = drr.GetInt32 (0);
				this.txtidProforma.Value = drr.GetInt32 (1);
				if (drr.GetValue(2) != System.DBNull.Value)this.txtIdCliente.Value= drr.GetInt32 (2);
				this.dtFecha.Value = drr.GetDateTime (3);
				if (drr.GetValue(4) != System.DBNull.Value)this.txtTitulo.Text= drr.GetString (4);
				if (drr.GetValue(5) != System.DBNull.Value)this.txtEstado.Value = drr.GetInt32 (5);
				if (drr.GetValue(6) != System.DBNull.Value)this.txtArchivo.Text= drr.GetString (6);
				if (drr.GetValue(7) != System.DBNull.Value)this.txtNombre.Text= drr.GetString (7);
				drr.Close();
				
				#endregion Consulta Individual

			}
			else
			{
				this.txtIdCliente.Value = idcliente;
				this.txtEstado.Value = iestado;
				this.txtidProforma.Value = iproforma;
				this.txtidcontrato.Value = idContrato;
				this.dtFecha.Focus();
				this.dtFecha.Value = DateTime.Today;
				this.dtFecha.CalendarInfo.MinDate = DateTime.Parse(dtfecha.ToString());
				this.txtNombre.Text = snombre;

			}

			if( (int) this.txtEstado.Value < 2)
			{
				this.lblEstado.Text ="PENDIENTE";
			  this.dtFecha.Enabled = true;
				this.txtArchivo.Enabled = true;
				this.txtTitulo.Enabled = true;

			}
			if( (int) this.txtEstado.Value == 2)
			{
				this.lblEstado.Text ="AUTORIZADO";
				this.dtFecha.Enabled = false;
				this.txtArchivo.Enabled = false;
				this.txtTitulo.Enabled = false;
				
			}
			if( (int) this.txtEstado.Value == 3)
			{
				this.lblEstado.Text ="NEGADO ";
				this.dtFecha.Enabled = false;
				this.txtArchivo.Enabled = false;
				this.txtTitulo.Enabled = false;
			}


		}

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				if (e.KeyCode == Keys.Enter) this.txtTitulo.Focus();
		}

		private void txtTitulo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				if (e.KeyCode == Keys.Enter) this.txtArchivo.Focus();
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Seleccione la Fecha Inicial", false, "", cdsSeteoF)) return;
			if (!Validacion.vbTexto(this.txtTitulo, 10, 1000, "El Titulo")) return;
			if (!Validacion.vbTexto(this.txtArchivo, 10, 1000, "La Nota")) return;
			this.DialogResult = DialogResult.OK;

		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;

		}
	}
}






