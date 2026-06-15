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
	/// Summary description for frmAnulaFacturas.
	/// </summary>
	public class frmAnulaFacturas : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDesde;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtHasta;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.GroupBox groupBox1;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmAnulaFacturas()
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
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmAnulaFacturas));
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			this.txtDesde = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtHasta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblBodega = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.txtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// txtDesde
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDesde.Appearance = appearance1;
			this.txtDesde.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDesde.FormatString = "";
			this.txtDesde.Location = new System.Drawing.Point(56, 80);
			this.txtDesde.MaskInput = "nnnnnnnnn";
			this.txtDesde.Name = "txtDesde";
			this.txtDesde.PromptChar = ' ';
			this.txtDesde.Size = new System.Drawing.Size(80, 22);
			this.txtDesde.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtDesde.TabIndex = 2;
			this.txtDesde.Click += new System.EventHandler(this.txtDesde_Click);
			this.txtDesde.ValueChanged += new System.EventHandler(this.txtDesde_ValueChanged);
			this.txtDesde.Enter += new System.EventHandler(this.txtDesde_Enter);
			// 
			// txtHasta
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtHasta.Appearance = appearance2;
			this.txtHasta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtHasta.FormatString = "";
			this.txtHasta.Location = new System.Drawing.Point(224, 80);
			this.txtHasta.MaskInput = "nnnnnnnnn";
			this.txtHasta.Name = "txtHasta";
			this.txtHasta.PromptChar = ' ';
			this.txtHasta.Size = new System.Drawing.Size(80, 22);
			this.txtHasta.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtHasta.TabIndex = 3;
			this.txtHasta.Click += new System.EventHandler(this.txtHasta_Click);
			this.txtHasta.Enter += new System.EventHandler(this.txtHasta_Enter);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(8, 80);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(35, 16);
			this.label1.TabIndex = 128;
			this.label1.Text = "Desde";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(168, 80);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(33, 16);
			this.label2.TabIndex = 129;
			this.label2.Text = "Hasta";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 48);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(34, 16);
			this.label6.TabIndex = 155;
			this.label6.Text = "Fecha";
			// 
			// dtFecha
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance3;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(56, 48);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 1;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.Validated += new System.EventHandler(this.dtFecha_Validated);
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(8, 16);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(30, 16);
			this.lblBodega.TabIndex = 162;
			this.lblBodega.Text = "Local";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 180;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(56, 16);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(248, 21);
			this.cmbBodega.TabIndex = 0;
			this.cmbBodega.ValueMember = "Bodega";
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
			// btnAnular
			// 
			this.btnAnular.Location = new System.Drawing.Point(8, 136);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.TabIndex = 4;
			this.btnAnular.Text = "&Anular";
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Location = new System.Drawing.Point(96, 136);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 5;
			this.btnSalir.Text = "Salir";
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 112);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(312, 8);
			this.groupBox1.TabIndex = 165;
			this.groupBox1.TabStop = false;
			// 
			// frmAnulaFacturas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(312, 166);
			this.ControlBox = false;
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtHasta);
			this.Controls.Add(this.txtDesde);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.Name = "frmAnulaFacturas";
			this.Text = "Anula Facturas";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmAnulaFacturas_KeyDown);
			this.Load += new System.EventHandler(this.frmAnulaFacturas_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmAnulaFacturas_Load(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Exec ListaBodegas {0}, 1", MenuLatinium.IdUsuario);
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;
			this.dtFecha.Value = DateTime.Today;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void frmAnulaFacturas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void dtFecha_Validated(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value == System.DBNull.Value) this.dtFecha.Value = DateTime.Today;
		}

		private void txtDesde_ValueChanged(object sender, System.EventArgs e)
		{
			this.txtHasta.MinValue = this.txtDesde.Value;
		}

		private void txtDesde_Enter(object sender, System.EventArgs e)
		{
			this.txtDesde.SelectAll();
		}

		private void txtDesde_Click(object sender, System.EventArgs e)
		{
			this.txtDesde.SelectAll();
		}

		private void txtHasta_Click(object sender, System.EventArgs e)
		{
			this.txtHasta.SelectAll();
		}

		private void txtHasta_Enter(object sender, System.EventArgs e)
		{
			this.txtHasta.SelectAll();
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{
			if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Venta"))
			{				
				DateTime dtFecha = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No Puede Anular Facturas en esta Fecha, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);								
			}
			else if (this.cmbBodega.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un Local", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);								
				this.cmbBodega.Focus();
			}
			else if (DialogResult.Yes == MessageBox.Show(string.Format("Esta seguro de Anular las Facturas del Local '{0}' desde {1} hasta {2}", this.cmbBodega.Text.ToString(), (int)this.txtDesde.Value, (int)this.txtHasta.Value), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
			{
				#region S.R.I.
				string sSerie = "";			
				DateTime dtFechaCaducidad = DateTime.Today;
				string sAutorizacion = "";
				string sNumero = "";
				DateTime dtFecha = (DateTime)this.dtFecha.Value;

				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, string.Format("Select Serie, Autorizacion, Caducidad, Automatico From Bodega Where Bodega = {0}", (int)this.cmbBodega.Value), true);	
				dr.Read();
				if (dr.HasRows)
				{
					sSerie = dr.GetValue(0).ToString();
					sAutorizacion = dr.GetValue(1).ToString();
					if (dr.GetValue(2) != System.DBNull.Value) dtFechaCaducidad = dr.GetDateTime(2);						
				}
				dr.Close();					
				#endregion S.R.I.

				this.Cursor = Cursors.WaitCursor;
				for (int i = (int)this.txtDesde.Value; i<= (int)this.txtHasta.Value; i++)
				{
					char Pad = '0';
					sNumero = i.ToString().PadLeft(9, Pad);

					int idCompra = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec GrabaMaestroFactura 0, 0, 0, 1, 15, 10, 0, 0, 1, {0}, 0, 0, '{1}', '', 0, 0, '', '', '{2}', '{3}', '{4}', '{5}', 0, '{2}', '{2}', 0, 'ANULACION POR CADUCIDAD', 0, 0, 0, 0, '{2}', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 0, 0, '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', 0, 0, '', '', 0, '', 0, 0, '', 0, 0, ''",
						(int)this.cmbBodega.Value, sNumero, dtFecha.ToString("yyyyMMdd"), sSerie, sAutorizacion, dtFechaCaducidad.ToString("yyyyMMdd")));

					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Insert Into DetalleFormaPago (idCompra, idFormaPago, Valor, Origen) Values ({0}, 1, 0, 1)", idCompra));
				}
				this.Cursor = Cursors.Default;

				MessageBox.Show("Facturas Anuladas Correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
		}
	}
}

