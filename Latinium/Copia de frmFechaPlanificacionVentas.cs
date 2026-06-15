using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmFechaPlanificacionVentas.
	/// </summary>
	public class frmFechaPlanificacionVentas : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.GroupBox groupBox1;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtObservacionesEntrada;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtObservacionesSalida;
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Button btnCancelar;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtpFecha;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHoraSalida;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHoraEntrada;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		private C1.Data.C1DataSet cdsSeteoF;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private System.Windows.Forms.Label lblEstado;

		DateTime dtFecha = DateTime.Today;
		int idDetallePlanificacionVentas = 0;

		public frmFechaPlanificacionVentas(DateTime DTFecha, int IdDetallePlanificacionVentas)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			dtFecha = DTFecha;
			idDetallePlanificacionVentas = IdDetallePlanificacionVentas;
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
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmFechaPlanificacionVentas));
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.txtObservacionesEntrada = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtObservacionesSalida = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.dtpFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtHoraSalida = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtHoraEntrada = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.txtObservacionesEntrada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservacionesSalida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHoraSalida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHoraEntrada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			this.SuspendLayout();
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Enabled = false;
			this.label4.Location = new System.Drawing.Point(8, 88);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(138, 16);
			this.label4.TabIndex = 289;
			this.label4.Text = "Observaciones de Entrada";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Enabled = false;
			this.label3.Location = new System.Drawing.Point(8, 200);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(130, 16);
			this.label3.TabIndex = 288;
			this.label3.Text = "Observaciones de Salida";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Enabled = false;
			this.label2.Location = new System.Drawing.Point(8, 10);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(36, 16);
			this.label2.TabIndex = 281;
			this.label2.Text = "Fecha";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Enabled = false;
			this.label1.Location = new System.Drawing.Point(8, 170);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(79, 16);
			this.label1.TabIndex = 279;
			this.label1.Text = "Hora de Salida";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Enabled = false;
			this.label6.Location = new System.Drawing.Point(8, 58);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(87, 16);
			this.label6.TabIndex = 278;
			this.label6.Text = "Hora de Entrada";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox2
			// 
			this.groupBox2.Location = new System.Drawing.Point(0, 280);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(624, 8);
			this.groupBox2.TabIndex = 287;
			this.groupBox2.TabStop = false;
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 40);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(624, 8);
			this.groupBox1.TabIndex = 286;
			this.groupBox1.TabStop = false;
			// 
			// txtObservacionesEntrada
			// 
			this.txtObservacionesEntrada.AcceptsReturn = true;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtObservacionesEntrada.Appearance = appearance1;
			this.txtObservacionesEntrada.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObservacionesEntrada.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObservacionesEntrada.Location = new System.Drawing.Point(8, 112);
			this.txtObservacionesEntrada.Multiline = true;
			this.txtObservacionesEntrada.Name = "txtObservacionesEntrada";
			this.txtObservacionesEntrada.Size = new System.Drawing.Size(600, 48);
			this.txtObservacionesEntrada.TabIndex = 285;
			// 
			// txtObservacionesSalida
			// 
			this.txtObservacionesSalida.AcceptsReturn = true;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtObservacionesSalida.Appearance = appearance2;
			this.txtObservacionesSalida.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObservacionesSalida.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObservacionesSalida.Location = new System.Drawing.Point(8, 224);
			this.txtObservacionesSalida.Multiline = true;
			this.txtObservacionesSalida.Name = "txtObservacionesSalida";
			this.txtObservacionesSalida.Size = new System.Drawing.Size(600, 48);
			this.txtObservacionesSalida.TabIndex = 284;
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.Location = new System.Drawing.Point(8, 296);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(88, 23);
			this.btnAceptar.TabIndex = 282;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.Location = new System.Drawing.Point(104, 296);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(88, 23);
			this.btnCancelar.TabIndex = 283;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// dtpFecha
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtpFecha.Appearance = appearance3;
			this.dtpFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtpFecha.DateButtons.Add(dateButton1);
			this.dtpFecha.Enabled = false;
			this.dtpFecha.Format = "dd/MM/yyyy";
			this.dtpFecha.Location = new System.Drawing.Point(56, 8);
			this.dtpFecha.Name = "dtpFecha";
			this.dtpFecha.NonAutoSizeHeight = 23;
			this.dtpFecha.Size = new System.Drawing.Size(112, 21);
			this.dtpFecha.SpinButtonsVisible = true;
			this.dtpFecha.TabIndex = 280;
			this.dtpFecha.Value = ((object)(resources.GetObject("dtpFecha.Value")));
			// 
			// dtHoraSalida
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHoraSalida.Appearance = appearance4;
			this.dtHoraSalida.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtHoraSalida.DateButtons.Add(dateButton2);
			this.dtHoraSalida.Enabled = false;
			this.dtHoraSalida.Format = "HH:mm";
			this.dtHoraSalida.Location = new System.Drawing.Point(104, 168);
			this.dtHoraSalida.Name = "dtHoraSalida";
			this.dtHoraSalida.NonAutoSizeHeight = 23;
			this.dtHoraSalida.Size = new System.Drawing.Size(112, 21);
			this.dtHoraSalida.SpinButtonsVisible = true;
			this.dtHoraSalida.TabIndex = 277;
			this.dtHoraSalida.Value = ((object)(resources.GetObject("dtHoraSalida.Value")));
			// 
			// dtHoraEntrada
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHoraEntrada.Appearance = appearance5;
			this.dtHoraEntrada.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.dtHoraEntrada.DateButtons.Add(dateButton3);
			this.dtHoraEntrada.Enabled = false;
			this.dtHoraEntrada.Format = "HH:mm";
			this.dtHoraEntrada.Location = new System.Drawing.Point(104, 56);
			this.dtHoraEntrada.Name = "dtHoraEntrada";
			this.dtHoraEntrada.NonAutoSizeHeight = 23;
			this.dtHoraEntrada.Size = new System.Drawing.Size(112, 21);
			this.dtHoraEntrada.SpinButtonsVisible = true;
			this.dtHoraEntrada.TabIndex = 276;
			this.dtHoraEntrada.Value = ((object)(resources.GetObject("dtHoraEntrada.Value")));
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
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(396, 296);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 21);
			this.txtEstado.TabIndex = 291;
			this.txtEstado.Visible = false;
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(204, 9);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 19);
			this.lblEstado.TabIndex = 290;
			// 
			// frmFechaPlanificacionVentas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(616, 326);
			this.ControlBox = false;
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtObservacionesEntrada);
			this.Controls.Add(this.txtObservacionesSalida);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.dtpFecha);
			this.Controls.Add(this.dtHoraSalida);
			this.Controls.Add(this.dtHoraEntrada);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.Name = "frmFechaPlanificacionVentas";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Horario de Entrada Salida de Locales";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmFechaPlanificacionVentas_KeyDown);
			this.Load += new System.EventHandler(this.frmFechaPlanificacionVentas_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtObservacionesEntrada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservacionesSalida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHoraSalida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHoraEntrada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		public int iDia = 0;
		public int idHorarioPlanificacionVentas = 0;

		private void frmFechaPlanificacionVentas_Load(object sender, System.EventArgs e)
		{
			this.dtpFecha.Value = dtFecha;

			DateTime dtFechaHoraActual = Funcion.dtEscalarSQL(cdsSeteoF, "Exec RetornaFechaServidor");

			iDia = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select DATEPART(D, '{0}')", dtFecha.ToString("yyyyMMdd")));

			string sSQL = string.Format("Exec ConsultaHorariosPlanificacionVentasPorDias {0}, {1}", iDia, idDetallePlanificacionVentas);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				idHorarioPlanificacionVentas = dr.GetInt32(0);
				if (dr.GetValue(1) != System.DBNull.Value) this.dtHoraEntrada.Value = dr.GetDateTime(1);
				if (dr.GetValue(2) != System.DBNull.Value) this.txtObservacionesEntrada.Text = dr.GetString(2);
				
				if (dr.GetValue(3) != System.DBNull.Value) this.dtHoraSalida.Value = dr.GetDateTime(3);
				else this.dtHoraSalida.Value = dtFechaHoraActual;//DateTime.Now;

				if (dr.GetValue(4) != System.DBNull.Value) this.txtObservacionesSalida.Text = dr.GetString(4);
				if (dr.GetValue(5) != System.DBNull.Value) this.txtEstado.Value = dr.GetInt32(5);
				if (dr.GetValue(6) != System.DBNull.Value) this.lblEstado.Text = dr.GetString(6); 
			}
			else
			{
				dr.Close();
				this.dtHoraEntrada.Value = dtFechaHoraActual;//DateTime.Now;
			}
			dr.Close();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void frmFechaPlanificacionVentas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			if (this.dtHoraEntrada.Value != System.DBNull.Value) this.txtEstado.Value = 1;
			if (this.dtHoraSalida.Value != System.DBNull.Value) this.txtEstado.Value = 2;

			this.DialogResult = DialogResult.OK;
		}
	}
}
