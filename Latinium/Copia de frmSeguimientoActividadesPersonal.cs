using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmSeguimientoActividadesPersonal.
	/// </summary>
	public class frmSeguimientoActividadesPersonal : System.Windows.Forms.Form
	{
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdActividadesPersonal;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Button btnCancelar;
		public System.Windows.Forms.Label label1;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcion;
		public System.Windows.Forms.Label label4;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtpFecha;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmSeguimientoActividadesPersonal()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		bool bSeguimFinalizar = false;
		DateTime dtFecha = DateTime.Today;

		public frmSeguimientoActividadesPersonal(bool BSeguimFinalizar, DateTime DTFecha)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			bSeguimFinalizar = BSeguimFinalizar;
			dtFecha = DTFecha;
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmSeguimientoActividadesPersonal));
			this.txtIdActividadesPersonal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.label1 = new System.Windows.Forms.Label();
			this.txtDescripcion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.dtpFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			((System.ComponentModel.ISupportInitialize)(this.txtIdActividadesPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpFecha)).BeginInit();
			this.SuspendLayout();
			// 
			// txtIdActividadesPersonal
			// 
			this.txtIdActividadesPersonal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdActividadesPersonal.Enabled = false;
			this.txtIdActividadesPersonal.FormatString = "";
			this.txtIdActividadesPersonal.Location = new System.Drawing.Point(592, 168);
			this.txtIdActividadesPersonal.Name = "txtIdActividadesPersonal";
			this.txtIdActividadesPersonal.PromptChar = ' ';
			this.txtIdActividadesPersonal.Size = new System.Drawing.Size(16, 21);
			this.txtIdActividadesPersonal.TabIndex = 369;
			this.txtIdActividadesPersonal.Visible = false;
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 152);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(624, 8);
			this.groupBox1.TabIndex = 368;
			this.groupBox1.TabStop = false;
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.Location = new System.Drawing.Point(8, 168);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(88, 23);
			this.btnAceptar.TabIndex = 2;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.Location = new System.Drawing.Point(104, 168);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(88, 23);
			this.btnCancelar.TabIndex = 3;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 40);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(80, 16);
			this.label1.TabIndex = 367;
			this.label1.Text = "Observaciones";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDescripcion
			// 
			this.txtDescripcion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDescripcion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescripcion.Location = new System.Drawing.Point(8, 64);
			this.txtDescripcion.MaxLength = 1000;
			this.txtDescripcion.Multiline = true;
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.Size = new System.Drawing.Size(600, 80);
			this.txtDescripcion.TabIndex = 1;
			this.txtDescripcion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDescripcion_KeyDown);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Enabled = false;
			this.label4.Location = new System.Drawing.Point(8, 10);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(36, 16);
			this.label4.TabIndex = 366;
			this.label4.Text = "Fecha";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtpFecha
			// 
			this.dtpFecha.AllowMonthSelection = true;
			this.dtpFecha.AutoSelectionUpdate = true;
			this.dtpFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtpFecha.DateButtons.Add(dateButton1);
			this.dtpFecha.DayDisplayStyle = Infragistics.Win.UltraWinSchedule.DayDisplayStyle.DayNumberAndImage;
			this.dtpFecha.DayOfWeekDisplayStyle = Infragistics.Win.UltraWinSchedule.DayOfWeekDisplayStyle.FirstRow;
			this.dtpFecha.Format = "dd/MM/yyyy";
			this.dtpFecha.Location = new System.Drawing.Point(56, 8);
			this.dtpFecha.MonthOrientation = Infragistics.Win.UltraWinSchedule.MonthOrientation.DownThenAcross;
			this.dtpFecha.Name = "dtpFecha";
			this.dtpFecha.NonAutoSizeHeight = 23;
			this.dtpFecha.Size = new System.Drawing.Size(112, 21);
			this.dtpFecha.SpinButtonsVisible = true;
			this.dtpFecha.TabIndex = 0;
			this.dtpFecha.Value = ((object)(resources.GetObject("dtpFecha.Value")));
			// 
			// frmSeguimientoActividadesPersonal
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(618, 200);
			this.ControlBox = false;
			this.Controls.Add(this.txtIdActividadesPersonal);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtDescripcion);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.dtpFecha);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.Name = "frmSeguimientoActividadesPersonal";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Seguimiento";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmSeguimientoActividadesPersonal_KeyDown);
			this.Load += new System.EventHandler(this.frmSeguimientoActividadesPersonal_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtIdActividadesPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpFecha)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmSeguimientoActividadesPersonal_Load(object sender, System.EventArgs e)
		{
			this.dtpFecha.CalendarInfo.MinDate = dtFecha;
			this.dtpFecha.CalendarInfo.MaxDate = DateTime.Today;

			this.dtpFecha.Value = DateTime.Today;

			if (bSeguimFinalizar) this.Text = "Finalizar";
		}

		private void frmSeguimientoActividadesPersonal_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbTexto(this.txtDescripcion, 20, 1000, "Observaciones")) return;

			this.DialogResult = DialogResult.OK;
		}

		private void txtDescripcion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnAceptar.Focus();
		}
	}
}
