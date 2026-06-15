using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmSubeVentas.
	/// </summary>
	public class frmSubeVentas : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label18;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaVenta;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValor;
		public DevExpress.XtraEditors.SimpleButton btVer;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAno;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMes;
		private System.Windows.Forms.Label label7;
		private C1.Data.C1DataSet cdsSeteo;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmSubeVentas()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmSubeVentas));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			this.label18 = new System.Windows.Forms.Label();
			this.cmbFechaVenta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.txtValor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btVer = new DevExpress.XtraEditors.SimpleButton();
			this.txtAno = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.txtMes = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.cdsSeteo = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaVenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAno)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			this.SuspendLayout();
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label18.Location = new System.Drawing.Point(45, 105);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(87, 16);
			this.label18.TabIndex = 754;
			this.label18.Text = "Día CALCULAR";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label18.Visible = false;
			// 
			// cmbFechaVenta
			// 
			this.cmbFechaVenta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.cmbFechaVenta.DateButtons.Add(dateButton1);
			this.cmbFechaVenta.Format = "dd/MM/yyyy";
			this.cmbFechaVenta.Location = new System.Drawing.Point(137, 103);
			this.cmbFechaVenta.Name = "cmbFechaVenta";
			this.cmbFechaVenta.NonAutoSizeHeight = 23;
			this.cmbFechaVenta.Size = new System.Drawing.Size(104, 21);
			this.cmbFechaVenta.SpinButtonsVisible = true;
			this.cmbFechaVenta.TabIndex = 753;
			this.cmbFechaVenta.Value = ((object)(resources.GetObject("cmbFechaVenta.Value")));
			this.cmbFechaVenta.Visible = false;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(14, 97);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(32, 16);
			this.label2.TabIndex = 752;
			this.label2.Text = "Valor";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label2.Visible = false;
			// 
			// txtValor
			// 
			this.txtValor.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValor.Appearance = appearance1;
			this.txtValor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValor.FormatString = "#,##0.00";
			this.txtValor.Location = new System.Drawing.Point(46, 93);
			this.txtValor.Name = "txtValor";
			this.txtValor.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValor.PromptChar = ' ';
			this.txtValor.Size = new System.Drawing.Size(88, 21);
			this.txtValor.TabIndex = 751;
			this.txtValor.Visible = false;
			// 
			// btVer
			// 
			this.btVer.Appearance.Options.UseTextOptions = true;
			this.btVer.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.btVer.Image = ((System.Drawing.Image)(resources.GetObject("btVer.Image")));
			this.btVer.ImageAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.btVer.ImageIndex = 4;
			this.btVer.Location = new System.Drawing.Point(102, 153);
			this.btVer.Name = "btVer";
			this.btVer.Size = new System.Drawing.Size(48, 40);
			this.btVer.TabIndex = 750;
			this.btVer.ToolTip = "Procesar información";
			this.btVer.Click += new System.EventHandler(this.btVer_Click);
			// 
			// txtAno
			// 
			this.txtAno.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAno.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtAno.Location = new System.Drawing.Point(190, 40);
			this.txtAno.MinValue = 2014;
			this.txtAno.Name = "txtAno";
			this.txtAno.PromptChar = ' ';
			this.txtAno.Size = new System.Drawing.Size(88, 21);
			this.txtAno.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtAno.TabIndex = 748;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(160, 42);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(25, 16);
			this.label1.TabIndex = 749;
			this.label1.Text = "Año";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtMes
			// 
			this.txtMes.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMes.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtMes.Location = new System.Drawing.Point(46, 40);
			this.txtMes.MinValue = 1;
			this.txtMes.Name = "txtMes";
			this.txtMes.PromptChar = ' ';
			this.txtMes.Size = new System.Drawing.Size(88, 21);
			this.txtMes.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtMes.TabIndex = 746;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.Location = new System.Drawing.Point(16, 42);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(26, 16);
			this.label7.TabIndex = 747;
			this.label7.Text = "Mes";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cdsSeteo
			// 
			this.cdsSeteo.BindingContextCtrl = this;
			this.cdsSeteo.DataLibrary = "LibContabilidad";
			this.cdsSeteo.DataLibraryUrl = "";
			this.cdsSeteo.DataSetDef = "dsSeteo";
			this.cdsSeteo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteo.Name = "cdsSeteo";
			this.cdsSeteo.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeteo.SchemaDef = null;
			this.cdsSeteo.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteo_BeforeFill);
			// 
			// frmSubeVentas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(292, 262);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.cmbFechaVenta);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtValor);
			this.Controls.Add(this.btVer);
			this.Controls.Add(this.txtAno);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtMes);
			this.Controls.Add(this.label7);
			this.Name = "frmSubeVentas";
			this.Text = "Sube Ventas por Mes";
			this.Load += new System.EventHandler(this.frmSubeVentas_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaVenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAno)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmSubeVentas_Load(object sender, System.EventArgs e)
		{
			this.txtAno.MaxValue = DateTime.Today.Year;
			this.txtAno.Value = DateTime.Today.Year;
			this.txtMes.MaxValue = DateTime.Today.Month;
			this.txtMes.Value = DateTime.Today.Month;
			this.cmbFechaVenta.Value = DateTime.Now;
		}

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteo.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void btVer_Click(object sender, System.EventArgs e)
		{
			/*Realizar calculo de ventas*/
			Cursor = Cursors.WaitCursor;

			/* Sube información de loteservicio a LTSERVICIO */
			string sSQL = string.Format("Exec LTServicio.dbo.CalcularVentas_ValorFijo_TotVenta {0},{1}", 
				Convert.ToInt32(this.txtMes.Value), 
				Convert.ToInt32(this.txtMes.Value));
			Funcion.EjecutaSQL(cdsSeteo, sSQL);

			DateTime FeCalculo = (DateTime)this.cmbFechaVenta.Value;

			//string sSQL = string.Format("Exec LTServicio.dbo.CalcularVentas_Deposito '{0}'", 
			//	FeCalculo.ToString("yyyyMMdd"));
			//Funcion.EjecutaSQL(cdsSeteo, sSQL);

			//			string sSQL2 = string.Format("Exec LTServicio.dbo.ContabilizarVenta {0},{1},{3}", 
			//				Convert.ToInt32(this.txtMes.Value), 
			//				Convert.ToInt32(this.txtMes.Value),
			//				Convert.ToDecimal(this.txtValor.Value));
			//			Funcion.dvProcedimiento(cdsSeteo, sSQL2);

			Cursor = Cursors.Default;
			MessageBox.Show("Proceso TERMINADO...!!!", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
		}
	}
}
