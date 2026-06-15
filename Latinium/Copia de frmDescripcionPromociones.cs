using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.IO;
using System.Drawing;
using System.Data.SqlClient;
using System.Data;
using System.Text;
using System.Threading;
using System.Collections.Specialized;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmDescripcionPromociones.
	/// </summary>
	public class frmDescripcionPromociones : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label lblCliente;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFin;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtInicio;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.OpenFileDialog openFileDialog1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidDescripcionPromocion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPromocion;
		private System.Windows.Forms.PictureBox pbImagen;
		private System.Windows.Forms.Button btnFinalizar;
		private System.Windows.Forms.Button btnImagen;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private System.Windows.Forms.Label lblImagen;
		private System.Windows.Forms.RichTextBox rtbDescripcion;
		private System.Windows.Forms.Button btnColor;
		private System.Windows.Forms.Button btnFuente;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmDescripcionPromociones()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmDescripcionPromociones));
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.lblCliente = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.txtPromocion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtFin = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtInicio = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.btnFinalizar = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.pbImagen = new System.Windows.Forms.PictureBox();
			this.lblEstado = new System.Windows.Forms.Label();
			this.btnImagen = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtidDescripcionPromocion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblImagen = new System.Windows.Forms.Label();
			this.rtbDescripcion = new System.Windows.Forms.RichTextBox();
			this.btnFuente = new System.Windows.Forms.Button();
			this.btnColor = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.txtPromocion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFin)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtInicio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidDescripcionPromocion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			this.SuspendLayout();
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Enabled = false;
			this.lblCliente.Location = new System.Drawing.Point(8, 40);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(57, 16);
			this.lblCliente.TabIndex = 211;
			this.lblCliente.Text = "Promoción";
			this.lblCliente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Enabled = false;
			this.label1.Location = new System.Drawing.Point(240, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(33, 16);
			this.label1.TabIndex = 210;
			this.label1.Text = "Hasta";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Enabled = false;
			this.label6.Location = new System.Drawing.Point(8, 8);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(35, 16);
			this.label6.TabIndex = 209;
			this.label6.Text = "Desde";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtPromocion
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPromocion.Appearance = appearance1;
			this.txtPromocion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtPromocion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPromocion.Enabled = false;
			this.txtPromocion.Location = new System.Drawing.Point(80, 40);
			this.txtPromocion.Name = "txtPromocion";
			this.txtPromocion.Size = new System.Drawing.Size(592, 22);
			this.txtPromocion.TabIndex = 208;
			this.txtPromocion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtPromocion_KeyDown);
			// 
			// dtFin
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFin.Appearance = appearance2;
			this.dtFin.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFin.DateButtons.Add(dateButton1);
			this.dtFin.Enabled = false;
			this.dtFin.Format = "dd/MM/yyyy HH:mm";
			this.dtFin.Location = new System.Drawing.Point(296, 8);
			this.dtFin.Name = "dtFin";
			this.dtFin.NonAutoSizeHeight = 23;
			this.dtFin.Size = new System.Drawing.Size(144, 21);
			this.dtFin.SpinButtonsVisible = true;
			this.dtFin.TabIndex = 207;
			this.dtFin.Value = ((object)(resources.GetObject("dtFin.Value")));
			this.dtFin.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFin_KeyDown);
			// 
			// dtInicio
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtInicio.Appearance = appearance3;
			this.dtInicio.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtInicio.DateButtons.Add(dateButton2);
			this.dtInicio.Enabled = false;
			this.dtInicio.Format = "dd/MM/yyyy HH:mm";
			this.dtInicio.Location = new System.Drawing.Point(80, 8);
			this.dtInicio.Name = "dtInicio";
			this.dtInicio.NonAutoSizeHeight = 23;
			this.dtInicio.Size = new System.Drawing.Size(144, 21);
			this.dtInicio.SpinButtonsVisible = true;
			this.dtInicio.TabIndex = 206;
			this.dtInicio.Value = ((object)(resources.GetObject("dtInicio.Value")));
			this.dtInicio.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtInicio_KeyDown);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Enabled = false;
			this.label2.Location = new System.Drawing.Point(8, 72);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(61, 16);
			this.label2.TabIndex = 213;
			this.label2.Text = "Descripción";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnFinalizar
			// 
			this.btnFinalizar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnFinalizar.CausesValidation = false;
			this.btnFinalizar.Enabled = false;
			this.btnFinalizar.Image = ((System.Drawing.Image)(resources.GetObject("btnFinalizar.Image")));
			this.btnFinalizar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnFinalizar.Location = new System.Drawing.Point(314, 592);
			this.btnFinalizar.Name = "btnFinalizar";
			this.btnFinalizar.Size = new System.Drawing.Size(83, 23);
			this.btnFinalizar.TabIndex = 219;
			this.btnFinalizar.Text = "&Finalizar";
			this.btnFinalizar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnFinalizar.Click += new System.EventHandler(this.btnFinalizar_Click);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 592);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 214;
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
			this.btnConsultar.Location = new System.Drawing.Point(81, 592);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 218;
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
			this.btnEditar.Location = new System.Drawing.Point(160, 592);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 217;
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
			this.btnGuardar.Location = new System.Drawing.Point(237, 592);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 215;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(398, 592);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 221;
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
			this.btnSalir.Location = new System.Drawing.Point(471, 592);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 222;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// pbImagen
			// 
			this.pbImagen.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.pbImagen.Enabled = false;
			this.pbImagen.Location = new System.Drawing.Point(688, 72);
			this.pbImagen.Name = "pbImagen";
			this.pbImagen.Size = new System.Drawing.Size(416, 496);
			this.pbImagen.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
			this.pbImagen.TabIndex = 223;
			this.pbImagen.TabStop = false;
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(840, 16);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 224;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnImagen
			// 
			this.btnImagen.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImagen.CausesValidation = false;
			this.btnImagen.Enabled = false;
			this.btnImagen.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImagen.Location = new System.Drawing.Point(688, 40);
			this.btnImagen.Name = "btnImagen";
			this.btnImagen.Size = new System.Drawing.Size(104, 23);
			this.btnImagen.TabIndex = 225;
			this.btnImagen.Text = "&Cargar Imagen";
			this.btnImagen.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImagen.Click += new System.EventHandler(this.btnImagen_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 576);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1112, 8);
			this.groupBox1.TabIndex = 227;
			this.groupBox1.TabStop = false;
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
			// txtidDescripcionPromocion
			// 
			this.txtidDescripcionPromocion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidDescripcionPromocion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidDescripcionPromocion.Enabled = false;
			this.txtidDescripcionPromocion.FormatString = "#,##0.00";
			this.txtidDescripcionPromocion.Location = new System.Drawing.Point(1088, 592);
			this.txtidDescripcionPromocion.Name = "txtidDescripcionPromocion";
			this.txtidDescripcionPromocion.PromptChar = ' ';
			this.txtidDescripcionPromocion.Size = new System.Drawing.Size(16, 22);
			this.txtidDescripcionPromocion.TabIndex = 301;
			this.txtidDescripcionPromocion.Visible = false;
			// 
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.FormatString = "#,##0.00";
			this.txtEstado.Location = new System.Drawing.Point(1072, 592);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 22);
			this.txtEstado.TabIndex = 302;
			this.txtEstado.Visible = false;
			// 
			// lblImagen
			// 
			this.lblImagen.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblImagen.AutoSize = true;
			this.lblImagen.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblImagen.ForeColor = System.Drawing.Color.Firebrick;
			this.lblImagen.Location = new System.Drawing.Point(592, 592);
			this.lblImagen.Name = "lblImagen";
			this.lblImagen.Size = new System.Drawing.Size(0, 26);
			this.lblImagen.TabIndex = 304;
			this.lblImagen.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.lblImagen.Visible = false;
			// 
			// rtbDescripcion
			// 
			this.rtbDescripcion.AcceptsTab = true;
			this.rtbDescripcion.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.rtbDescripcion.Location = new System.Drawing.Point(8, 72);
			this.rtbDescripcion.Name = "rtbDescripcion";
			this.rtbDescripcion.Size = new System.Drawing.Size(664, 496);
			this.rtbDescripcion.TabIndex = 305;
			this.rtbDescripcion.Text = "";
			// 
			// btnFuente
			// 
			this.btnFuente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnFuente.CausesValidation = false;
			this.btnFuente.Enabled = false;
			this.btnFuente.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnFuente.Location = new System.Drawing.Point(448, 8);
			this.btnFuente.Name = "btnFuente";
			this.btnFuente.Size = new System.Drawing.Size(56, 23);
			this.btnFuente.TabIndex = 306;
			this.btnFuente.Text = "&Fuente";
			this.btnFuente.Click += new System.EventHandler(this.btnFuente_Click);
			// 
			// btnColor
			// 
			this.btnColor.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnColor.CausesValidation = false;
			this.btnColor.Enabled = false;
			this.btnColor.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnColor.Location = new System.Drawing.Point(512, 8);
			this.btnColor.Name = "btnColor";
			this.btnColor.Size = new System.Drawing.Size(56, 23);
			this.btnColor.TabIndex = 307;
			this.btnColor.Text = "&Color";
			this.btnColor.Click += new System.EventHandler(this.btnColor_Click);
			// 
			// frmDescripcionPromociones
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1112, 622);
			this.Controls.Add(this.btnColor);
			this.Controls.Add(this.btnFuente);
			this.Controls.Add(this.rtbDescripcion);
			this.Controls.Add(this.lblImagen);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtidDescripcionPromocion);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnImagen);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.pbImagen);
			this.Controls.Add(this.btnFinalizar);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.txtPromocion);
			this.Controls.Add(this.dtFin);
			this.Controls.Add(this.dtInicio);
			this.Name = "frmDescripcionPromociones";
			this.Text = "Promociones Activas";
			this.Load += new System.EventHandler(this.frmDescripcionPromociones_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtPromocion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFin)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtInicio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidDescripcionPromocion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		bool bNuevo = false;
		bool bEdicion = false;

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void frmDescripcionPromociones_Load(object sender, System.EventArgs e)
		{
		
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;

			this.txtidDescripcionPromocion.Value = 0;
			this.txtEstado.Value = 0;
			this.dtInicio.Value = System.DBNull.Value;
			this.dtFin.Value = System.DBNull.Value;
			this.txtPromocion.Text = "";
			this.rtbDescripcion.Text = "";
			this.pbImagen.Image = null;
			this.txtEstado.Value = 0;
			this.lblImagen.Text = "";

			this.dtInicio.Enabled = false;
			this.dtFin.Enabled = false;
			this.txtPromocion.Enabled = false;
			this.rtbDescripcion.Enabled = false;
			this.pbImagen.Enabled = false;

			this.lblEstado.Text = "";

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnFinalizar.Enabled = false;			
			this.btnCancelar.Enabled = false;
			this.btnImagen.Enabled = false;
			this.btnFuente.Enabled = false;
			this.btnColor.Enabled = false;

			this.btnFinalizar.Text = "Finalizar";
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			this.txtEstado.Value = 1;

			this.dtInicio.Enabled = true;
			this.dtFin.Enabled = true;
			this.txtPromocion.Enabled = true;			
			this.rtbDescripcion.Enabled = true;
			this.pbImagen.Enabled = true;

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnFinalizar.Enabled = false;			
			this.btnCancelar.Enabled = true;
			this.btnImagen.Enabled = true;
			this.btnFuente.Enabled = true;
			this.btnColor.Enabled = true;

			bNuevo = true;
			bEdicion = true;

			this.dtInicio.Focus();
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			this.dtInicio.Enabled = true;
			this.dtFin.Enabled = true;
			this.txtPromocion.Enabled = true;
			this.rtbDescripcion.Enabled = true;
			this.pbImagen.Enabled = true;

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnFinalizar.Enabled = false;			
			this.btnCancelar.Enabled = true;
			this.btnImagen.Enabled = true;

			this.btnFuente.Enabled = true;
			this.btnColor.Enabled = true;

			bNuevo = false;
			bEdicion = true;

			this.dtInicio.Focus();
		}

		private void btnFinalizar_Click(object sender, System.EventArgs e)
		{
			if ((int)this.txtEstado.Value == 1)
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DescripcionPromociones Set Estado = 2 Where idDescripcionPromociones = {0}", (int)this.txtidDescripcionPromocion.Value));
				this.lblEstado.Text = "FINALIZADO";
				this.btnFinalizar.Text = "Activar";
			}
			else if ((int)this.txtEstado.Value == 2)
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DescripcionPromociones Set Estado = 1 Where idDescripcionPromociones = {0}", (int)this.txtidDescripcionPromocion.Value));
				this.lblEstado.Text = "ACTIVO";
				this.btnFinalizar.Text = "Finalizar";
			}
		}

		private void dtInicio_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.dtFin.Focus();
		}

		private void dtFin_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtPromocion.Focus();
		}

		private void txtPromocion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.rtbDescripcion.Focus();
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if (this.dtInicio.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha de Inicio", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtInicio.Focus();
			}	
			else if (this.dtFin.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha de Inicio", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFin.Focus();
			}
			else if (this.txtPromocion.Text.Length == 0)
			{
				MessageBox.Show("Escriba el Nombre de la Promocion", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtPromocion.Focus();
			}
			else if (this.rtbDescripcion.Text.Length == 0)
			{
				MessageBox.Show("Escriba la Descripcion de la Promocion", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.rtbDescripcion.Focus();
			}
			else if (this.pbImagen.Image == null)
			{
				MessageBox.Show("Seleccione una Imagen", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtPromocion.Focus();
			}
			else
			{
				using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
				{
					oConexion.Open();

					SqlTransaction oTransaction = null;
					SqlCommand oCmdActualiza = oConexion.CreateCommand();
					oCmdActualiza.Connection = oConexion;
					oCmdActualiza.CommandType = CommandType.Text;
					oCmdActualiza.CommandTimeout = 0;
						
					try
					{						
						oTransaction = oConexion.BeginTransaction();
						oCmdActualiza.Transaction = oTransaction;
						
							MemoryStream ms = new MemoryStream();
							FileStream fs = new FileStream(this.lblImagen.Text, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
 
							ms.SetLength(fs.Length);
							fs.Read(ms.GetBuffer(), 0, (int)fs.Length);
 
							byte[] arrImg = ms.GetBuffer();
							ms.Flush();
							fs.Close();
						
						oCmdActualiza.CommandType = CommandType.StoredProcedure;
						oCmdActualiza.CommandText = "GuardaDescripcionPromociones";
						
						oCmdActualiza.Parameters.Add("@idDescripcionPromociones", SqlDbType.Int).Value = (int)this.txtidDescripcionPromocion.Value;
						oCmdActualiza.Parameters.Add("@Nombre", SqlDbType.VarChar).Value = this.txtPromocion.Text.ToString();						
						oCmdActualiza.Parameters.Add("@Desde", SqlDbType.DateTime).Value = (DateTime)this.dtInicio.Value;
						oCmdActualiza.Parameters.Add("@Hasta", SqlDbType.DateTime).Value = (DateTime)this.dtFin.Value;
						oCmdActualiza.Parameters.Add("@Descripcion", SqlDbType.VarChar).Value = this.rtbDescripcion.Text.ToString();
						oCmdActualiza.Parameters.Add("@Imagen", SqlDbType.VarBinary).Value = arrImg;
						
						this.txtidDescripcionPromocion.Value = oCmdActualiza.ExecuteScalar();

						oTransaction.Commit();

						this.dtInicio.Enabled = false;
						this.dtFin.Enabled = false;
						this.txtPromocion.Enabled = false;
						this.rtbDescripcion.Enabled = false;
						this.pbImagen.Enabled = false;

						this.btnNuevo.Enabled = true;
						this.btnConsultar.Enabled = true;
						this.btnEditar.Enabled = true;
						this.btnGuardar.Enabled = false;
						this.btnFinalizar.Enabled = true;			
						this.btnCancelar.Enabled = true;
						this.btnImagen.Enabled = false;

						this.btnFuente.Enabled = false;
						this.btnColor.Enabled = false;

						bNuevo = false;
						bEdicion = false;
					}
					catch (Exception ex)
					{
						try
						{
							oTransaction.Rollback();

							MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
						catch (Exception ex2)
						{
							MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
						}
					}
					finally
					{
						oConexion.Close();
					}
				}				
			}
		}

		private void btnImagen_Click(object sender, System.EventArgs e)
		{			
			this.openFileDialog1.CheckFileExists = true;
			this.openFileDialog1.ShowReadOnly = false;
			this.openFileDialog1.Filter = "All Files|*.*|Bitmap Files (*)|*.bmp;*.gif;*.jpg";
			this.openFileDialog1.FilterIndex = 2;
			
			if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
			{
				this.pbImagen.Image = Image.FromFile(this.openFileDialog1.FileName);
				this.lblImagen.Text = this.openFileDialog1.FileName;
			}			 
		}

		private void Consultar(int idDescripcion)
		{
			string sSQL = string.Format("Exec BuscaDescripcionPromociones '', {0}", idDescripcion);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtidDescripcionPromocion.Value = dr.GetInt32(0);
				this.txtPromocion.Text = dr.GetString(1);
				this.dtInicio.Value = dr.GetDateTime(2);
				this.dtFin.Value = dr.GetDateTime(3);
				//this.txtDescripcion.Text = dr.GetString(4);
				this.rtbDescripcion.Text = dr.GetString(4);
				byte[] arrImg = (byte[])dr.GetValue(5);
				this.txtEstado.Value = (int)dr.GetInt32(6);
				this.lblEstado.Text = dr.GetString(7);

				MemoryStream ms = new MemoryStream(arrImg);
				this.pbImagen.Image = System.Drawing.Image.FromStream(ms);
				ms.Close();
			}
			dr.Close();

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			if ((int)this.txtEstado.Value < 2) this.btnEditar.Enabled = true;
			this.btnGuardar.Enabled = false;
			this.btnFinalizar.Enabled = true;
			
			if ((int)this.txtEstado.Value == 2) this.btnFinalizar.Text = "Activar";

			this.btnCancelar.Enabled = true;
			this.btnImagen.Enabled = false;
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			btnCancelar_Click(sender, e);

			using (frmBuscaDescripciones Busqueda = new frmBuscaDescripciones())
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					try
					{	
						Consultar((int)Busqueda.ultraGrid1.ActiveRow.Cells["idDescripcionPromociones"].Value);
					}
					catch(Exception Exc)
					{
						MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}	
				}
			}
		}

		private void btnColor_Click(object sender, System.EventArgs e)
		{
			ColorDialog color = new ColorDialog();
			if (color.ShowDialog() == DialogResult.OK)
			{
				rtbDescripcion.ForeColor = color.Color;
			}
		}

		private void btnFuente_Click(object sender, System.EventArgs e)
		{
			FontDialog font = new FontDialog();
			font.Font = rtbDescripcion.Font;
			if (font.ShowDialog() == DialogResult.OK)
			{
				rtbDescripcion.Font = font.Font;
			}
		}
	}
}

