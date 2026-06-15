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
	/// Descripción breve de frmPlanificacionVentasVistaDiaria.
	/// </summary>
	public class frmPlanificacionVentasVistaDiaria : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.GroupBox groupBox1;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtObservacionesEntrada;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtObservacionesSalida;
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Button btnCancelar;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHoraSalida;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHoraEntrada;
		private System.Windows.Forms.Label lblSubTotal;
		private System.Windows.Forms.Label lblIva;
		private System.Windows.Forms.Label lblDescuento;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtResultado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMetaDiaria;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtVendido;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label lblEstado;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmPlanificacionVentasVistaDiaria()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		DateTime dtFecha = DateTime.Today;
		int idDetallePlanificacionVentas = 0;
		int iBodega = 0;

		public frmPlanificacionVentasVistaDiaria(DateTime DTFecha, int IdDetallePlanificacionVentas, int IBodega)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			dtFecha = DTFecha;
			idDetallePlanificacionVentas = IdDetallePlanificacionVentas;
			iBodega = IBodega;
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmPlanificacionVentasVistaDiaria));
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.txtObservacionesEntrada = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtObservacionesSalida = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.dtHoraSalida = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtHoraEntrada = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblSubTotal = new System.Windows.Forms.Label();
			this.lblIva = new System.Windows.Forms.Label();
			this.lblDescuento = new System.Windows.Forms.Label();
			this.txtResultado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtMetaDiaria = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtVendido = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.lblEstado = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.txtObservacionesEntrada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservacionesSalida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHoraSalida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHoraEntrada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtResultado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMetaDiaria)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVendido)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Enabled = false;
			this.label4.Location = new System.Drawing.Point(8, 40);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(134, 17);
			this.label4.TabIndex = 301;
			this.label4.Text = "Observaciones de Entrada";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Enabled = false;
			this.label3.Location = new System.Drawing.Point(8, 152);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(125, 17);
			this.label3.TabIndex = 300;
			this.label3.Text = "Observaciones de Salida";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Enabled = false;
			this.label1.Location = new System.Drawing.Point(8, 120);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(77, 17);
			this.label1.TabIndex = 293;
			this.label1.Text = "Hora de Salida";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Enabled = false;
			this.label6.Location = new System.Drawing.Point(8, 8);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(86, 17);
			this.label6.TabIndex = 292;
			this.label6.Text = "Hora de Entrada";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 312);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(624, 8);
			this.groupBox1.TabIndex = 298;
			this.groupBox1.TabStop = false;
			// 
			// txtObservacionesEntrada
			// 
			this.txtObservacionesEntrada.AcceptsReturn = true;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtObservacionesEntrada.Appearance = appearance1;
			this.txtObservacionesEntrada.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObservacionesEntrada.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObservacionesEntrada.Enabled = false;
			this.txtObservacionesEntrada.Location = new System.Drawing.Point(8, 64);
			this.txtObservacionesEntrada.Multiline = true;
			this.txtObservacionesEntrada.Name = "txtObservacionesEntrada";
			this.txtObservacionesEntrada.Size = new System.Drawing.Size(600, 48);
			this.txtObservacionesEntrada.TabIndex = 297;
			// 
			// txtObservacionesSalida
			// 
			this.txtObservacionesSalida.AcceptsReturn = true;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtObservacionesSalida.Appearance = appearance2;
			this.txtObservacionesSalida.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObservacionesSalida.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObservacionesSalida.Enabled = false;
			this.txtObservacionesSalida.Location = new System.Drawing.Point(8, 176);
			this.txtObservacionesSalida.Multiline = true;
			this.txtObservacionesSalida.Name = "txtObservacionesSalida";
			this.txtObservacionesSalida.Size = new System.Drawing.Size(600, 48);
			this.txtObservacionesSalida.TabIndex = 296;
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.Location = new System.Drawing.Point(8, 324);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(88, 23);
			this.btnAceptar.TabIndex = 294;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.Location = new System.Drawing.Point(104, 324);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(88, 23);
			this.btnCancelar.TabIndex = 295;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// dtHoraSalida
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHoraSalida.Appearance = appearance3;
			this.dtHoraSalida.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHoraSalida.DateButtons.Add(dateButton1);
			this.dtHoraSalida.Enabled = false;
			this.dtHoraSalida.Format = "HH:mm";
			this.dtHoraSalida.Location = new System.Drawing.Point(104, 120);
			this.dtHoraSalida.Name = "dtHoraSalida";
			this.dtHoraSalida.NonAutoSizeHeight = 23;
			this.dtHoraSalida.Size = new System.Drawing.Size(112, 21);
			this.dtHoraSalida.SpinButtonsVisible = true;
			this.dtHoraSalida.TabIndex = 291;
			this.dtHoraSalida.Value = ((object)(resources.GetObject("dtHoraSalida.Value")));
			// 
			// dtHoraEntrada
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHoraEntrada.Appearance = appearance4;
			this.dtHoraEntrada.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtHoraEntrada.DateButtons.Add(dateButton2);
			this.dtHoraEntrada.Enabled = false;
			this.dtHoraEntrada.Format = "HH:mm";
			this.dtHoraEntrada.Location = new System.Drawing.Point(104, 8);
			this.dtHoraEntrada.Name = "dtHoraEntrada";
			this.dtHoraEntrada.NonAutoSizeHeight = 23;
			this.dtHoraEntrada.Size = new System.Drawing.Size(112, 21);
			this.dtHoraEntrada.SpinButtonsVisible = true;
			this.dtHoraEntrada.TabIndex = 290;
			this.dtHoraEntrada.Value = ((object)(resources.GetObject("dtHoraEntrada.Value")));
			// 
			// lblSubTotal
			// 
			this.lblSubTotal.AutoSize = true;
			this.lblSubTotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblSubTotal.Location = new System.Drawing.Point(8, 232);
			this.lblSubTotal.Name = "lblSubTotal";
			this.lblSubTotal.Size = new System.Drawing.Size(46, 16);
			this.lblSubTotal.TabIndex = 304;
			this.lblSubTotal.Text = "Vendido";
			this.lblSubTotal.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblIva
			// 
			this.lblIva.AutoSize = true;
			this.lblIva.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblIva.Location = new System.Drawing.Point(8, 280);
			this.lblIva.Name = "lblIva";
			this.lblIva.Size = new System.Drawing.Size(55, 16);
			this.lblIva.TabIndex = 303;
			this.lblIva.Text = "Resultado";
			this.lblIva.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblDescuento
			// 
			this.lblDescuento.AutoSize = true;
			this.lblDescuento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblDescuento.Location = new System.Drawing.Point(8, 256);
			this.lblDescuento.Name = "lblDescuento";
			this.lblDescuento.Size = new System.Drawing.Size(62, 16);
			this.lblDescuento.TabIndex = 302;
			this.lblDescuento.Text = "Meta Diaria";
			this.lblDescuento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtResultado
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtResultado.Appearance = appearance5;
			this.txtResultado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtResultado.Enabled = false;
			this.txtResultado.FormatString = "#,##0.00";
			this.txtResultado.Location = new System.Drawing.Point(192, 280);
			this.txtResultado.Name = "txtResultado";
			this.txtResultado.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtResultado.PromptChar = ' ';
			this.txtResultado.Size = new System.Drawing.Size(104, 22);
			this.txtResultado.TabIndex = 307;
			// 
			// txtMetaDiaria
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMetaDiaria.Appearance = appearance6;
			this.txtMetaDiaria.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMetaDiaria.Enabled = false;
			this.txtMetaDiaria.FormatString = "#,##0.00";
			this.txtMetaDiaria.Location = new System.Drawing.Point(192, 256);
			this.txtMetaDiaria.Name = "txtMetaDiaria";
			this.txtMetaDiaria.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMetaDiaria.PromptChar = ' ';
			this.txtMetaDiaria.Size = new System.Drawing.Size(104, 22);
			this.txtMetaDiaria.TabIndex = 306;
			// 
			// txtVendido
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVendido.Appearance = appearance7;
			this.txtVendido.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVendido.Enabled = false;
			this.txtVendido.FormatString = "#,##0.00";
			this.txtVendido.Location = new System.Drawing.Point(192, 232);
			this.txtVendido.Name = "txtVendido";
			this.txtVendido.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtVendido.PromptChar = ' ';
			this.txtVendido.Size = new System.Drawing.Size(104, 22);
			this.txtVendido.TabIndex = 305;
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
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(312, 280);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 19);
			this.lblEstado.TabIndex = 308;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// frmPlanificacionVentasVistaDiaria
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(614, 352);
			this.ControlBox = false;
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.lblSubTotal);
			this.Controls.Add(this.lblIva);
			this.Controls.Add(this.lblDescuento);
			this.Controls.Add(this.txtResultado);
			this.Controls.Add(this.txtMetaDiaria);
			this.Controls.Add(this.txtVendido);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtObservacionesEntrada);
			this.Controls.Add(this.txtObservacionesSalida);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.dtHoraSalida);
			this.Controls.Add(this.dtHoraEntrada);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.Name = "frmPlanificacionVentasVistaDiaria";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Resumen Planificacion Jefe de Ventas";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmPlanificacionVentasVistaDiaria_KeyDown);
			this.Load += new System.EventHandler(this.frmPlanificacionVentasVistaDiaria_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtObservacionesEntrada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservacionesSalida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHoraSalida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHoraEntrada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtResultado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMetaDiaria)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVendido)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmPlanificacionVentasVistaDiaria_Load(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Exec ResumenDiarioPlanificacionVentas {0}, {1}, '{2}'", 
				iBodega, idDetallePlanificacionVentas, dtFecha.ToString("yyyyMMdd"));
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				if (dr.GetValue(0) != System.DBNull.Value) this.dtHoraEntrada.Value = dr.GetDateTime(0);
				if (dr.GetValue(1) != System.DBNull.Value) this.txtObservacionesEntrada.Text = dr.GetString(1);
				
				if (dr.GetValue(2) != System.DBNull.Value) this.dtHoraSalida.Value = dr.GetDateTime(2);
				if (dr.GetValue(3) != System.DBNull.Value) this.txtObservacionesSalida.Text = dr.GetString(3);

				if (dr.GetValue(4) != System.DBNull.Value) this.txtVendido.Value = dr.GetDecimal(4);
				if (dr.GetValue(5) != System.DBNull.Value) this.txtMetaDiaria.Value = dr.GetDecimal(5);
				if (dr.GetValue(6) != System.DBNull.Value) this.txtResultado.Value = dr.GetDecimal(6);
				if (dr.GetValue(7) != System.DBNull.Value) this.lblEstado.Text = dr.GetString(7); 
			}
			dr.Close();
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

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.OK;
		}

		private void frmPlanificacionVentasVistaDiaria_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}
	}
}
