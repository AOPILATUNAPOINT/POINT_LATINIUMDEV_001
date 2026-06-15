using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Globalization;
using System.Data;
using System.Data.SqlClient;
using C1.Data;
using Infragistics.Win.UltraWinGrid;
using System.Web.Mail;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de IntCostoIndirecto.
	/// </summary>
	public class IntCostoIndirecto : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.Misc.UltraButton btnEgreso_CI;
		private Infragistics.Win.Misc.UltraButton btnCompra_CI;
		private Infragistics.Win.Misc.UltraButton btnAsiento_CI;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private C1.Data.C1DataSet cdsAsiento;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaH;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaD;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtparametro;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		private Infragistics.Win.Misc.UltraButton ultraButton2;
		private Infragistics.Win.Misc.UltraButton ultraButton3;
		private System.Windows.Forms.Label label4;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public IntCostoIndirecto()
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.btnEgreso_CI = new Infragistics.Win.Misc.UltraButton();
			this.btnCompra_CI = new Infragistics.Win.Misc.UltraButton();
			this.btnAsiento_CI = new Infragistics.Win.Misc.UltraButton();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbFechaH = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbFechaD = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.cdsAsiento = new C1.Data.C1DataSet();
			this.txtparametro = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			this.ultraButton2 = new Infragistics.Win.Misc.UltraButton();
			this.ultraButton3 = new Infragistics.Win.Misc.UltraButton();
			this.label4 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaH)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaD)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtparametro)).BeginInit();
			this.SuspendLayout();
			// 
			// btnEgreso_CI
			// 
			this.btnEgreso_CI.Location = new System.Drawing.Point(176, 112);
			this.btnEgreso_CI.Name = "btnEgreso_CI";
			this.btnEgreso_CI.Size = new System.Drawing.Size(80, 24);
			this.btnEgreso_CI.TabIndex = 0;
			this.btnEgreso_CI.Text = "&Egresos";
			this.btnEgreso_CI.Click += new System.EventHandler(this.btnEgreso_CI_Click);
			// 
			// btnCompra_CI
			// 
			this.btnCompra_CI.Location = new System.Drawing.Point(176, 152);
			this.btnCompra_CI.Name = "btnCompra_CI";
			this.btnCompra_CI.Size = new System.Drawing.Size(80, 24);
			this.btnCompra_CI.TabIndex = 1;
			this.btnCompra_CI.Text = "&Compras";
			this.btnCompra_CI.Click += new System.EventHandler(this.btnCompra_CI_Click);
			// 
			// btnAsiento_CI
			// 
			this.btnAsiento_CI.Location = new System.Drawing.Point(176, 192);
			this.btnAsiento_CI.Name = "btnAsiento_CI";
			this.btnAsiento_CI.Size = new System.Drawing.Size(80, 24);
			this.btnAsiento_CI.TabIndex = 3;
			this.btnAsiento_CI.Text = "&Asientos";
			this.btnAsiento_CI.Click += new System.EventHandler(this.btnAsiento_CI_Click);
			// 
			// label1
			// 
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(112, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(248, 24);
			this.label1.TabIndex = 4;
			this.label1.Text = "Envio de Costos Indirectos";
			// 
			// cmbFechaH
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFechaH.Appearance = appearance1;
			this.cmbFechaH.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.cmbFechaH.DateButtons.Add(dateButton1);
			this.cmbFechaH.Format = "dd/MMM/yyyy";
			this.cmbFechaH.Location = new System.Drawing.Point(328, 64);
			this.cmbFechaH.Name = "cmbFechaH";
			this.cmbFechaH.NonAutoSizeHeight = 23;
			this.cmbFechaH.Size = new System.Drawing.Size(112, 21);
			this.cmbFechaH.SpinButtonsVisible = true;
			this.cmbFechaH.TabIndex = 34;
			// 
			// cmbFechaD
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFechaD.Appearance = appearance2;
			this.cmbFechaD.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.cmbFechaD.DateButtons.Add(dateButton2);
			this.cmbFechaD.Format = "dd/MMM/yyyy";
			this.cmbFechaD.Location = new System.Drawing.Point(104, 64);
			this.cmbFechaD.Name = "cmbFechaD";
			this.cmbFechaD.NonAutoSizeHeight = 23;
			this.cmbFechaD.Size = new System.Drawing.Size(120, 21);
			this.cmbFechaD.SpinButtonsVisible = true;
			this.cmbFechaD.TabIndex = 33;
			// 
			// label2
			// 
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(16, 64);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(80, 24);
			this.label2.TabIndex = 35;
			this.label2.Text = "Fecha Desde";
			// 
			// label3
			// 
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.Location = new System.Drawing.Point(240, 64);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(88, 24);
			this.label3.TabIndex = 36;
			this.label3.Text = "Fecha Hasta";
			// 
			// cdsAsiento
			// 
			this.cdsAsiento.BindingContextCtrl = this;
			this.cdsAsiento.DataLibrary = "LibContabilidad";
			this.cdsAsiento.DataLibraryUrl = "";
			this.cdsAsiento.DataSetDef = "dsAsiento";
			this.cdsAsiento.FillOnRequest = false;
			this.cdsAsiento.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsAsiento.Name = "cdsAsiento";
			this.cdsAsiento.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsAsiento.SchemaDef = null;
			// 
			// txtparametro
			// 
			this.txtparametro.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtparametro.Location = new System.Drawing.Point(16, 152);
			this.txtparametro.Name = "txtparametro";
			this.txtparametro.Size = new System.Drawing.Size(128, 21);
			this.txtparametro.TabIndex = 101;
			// 
			// ultraButton1
			// 
			this.ultraButton1.Location = new System.Drawing.Point(280, 192);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.Size = new System.Drawing.Size(160, 24);
			this.ultraButton1.TabIndex = 104;
			this.ultraButton1.Text = "Habilitar Reenvio Asientos";
			this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// ultraButton2
			// 
			this.ultraButton2.Location = new System.Drawing.Point(280, 152);
			this.ultraButton2.Name = "ultraButton2";
			this.ultraButton2.Size = new System.Drawing.Size(160, 24);
			this.ultraButton2.TabIndex = 103;
			this.ultraButton2.Text = "Habilitar Reenvio  Compras";
			this.ultraButton2.Click += new System.EventHandler(this.ultraButton2_Click);
			// 
			// ultraButton3
			// 
			this.ultraButton3.Location = new System.Drawing.Point(280, 112);
			this.ultraButton3.Name = "ultraButton3";
			this.ultraButton3.Size = new System.Drawing.Size(160, 24);
			this.ultraButton3.TabIndex = 102;
			this.ultraButton3.Text = "Habilitar Reenvio Egresos";
			this.ultraButton3.Click += new System.EventHandler(this.ultraButton3_Click);
			// 
			// label4
			// 
			this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label4.Location = new System.Drawing.Point(16, 112);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(128, 24);
			this.label4.TabIndex = 105;
			this.label4.Text = "Código de Referencia";
			this.label4.Click += new System.EventHandler(this.label4_Click);
			// 
			// IntCostoIndirecto
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(464, 229);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.ultraButton1);
			this.Controls.Add(this.ultraButton2);
			this.Controls.Add(this.ultraButton3);
			this.Controls.Add(this.txtparametro);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.cmbFechaH);
			this.Controls.Add(this.cmbFechaD);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.btnAsiento_CI);
			this.Controls.Add(this.btnCompra_CI);
			this.Controls.Add(this.btnEgreso_CI);
			this.MaximizeBox = false;
			this.Name = "IntCostoIndirecto";
			this.Text = "Interfaz Costo Indirecto";
			this.Load += new System.EventHandler(this.IntCostoIndirecto_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaH)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaD)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtparametro)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		
		private void btnEgreso_CI_Click(object sender, System.EventArgs e)
		{
			if (MessageBox.Show("¿Desea Enviar Egresos?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) 
		{
			return;
		}
		else
			{
				enviaCostosIndirectos("Egreso_CI");		
			}

		}
		private void enviaCostosIndirectos(string tipo)
		{
			DateTime dtHasta;
			DateTime dtDesde;
			try
			{
				dtDesde = (DateTime) cmbFechaD.Value;
				dtHasta = (DateTime) cmbFechaH.Value;
			}
			catch(Exception ex)
			{
				MessageBox.Show("Error: "+ex);
				return;
			}
			Random rNumero = new Random();
			int iNumero = rNumero.Next(10000000);
			txtparametro.Text = iNumero.ToString();

			string stExec = "Exec Agro_MovimientoCI '"+tipo+"', '"				
				+dtDesde.ToString("yyyyMMdd")+"', '"
				+dtHasta.ToString("yyyyMMdd")+"','"
				+iNumero +"'";
			
			string mesj = Funcion.sEscalarSQL(cdsAsiento, stExec);

			if (mesj.Length > 0)
			{
				MessageBox.Show(mesj,"Información",MessageBoxButtons.OK,MessageBoxIcon.Information);
			}
		}

		private void btnCompra_CI_Click(object sender, System.EventArgs e)
		{
			if (MessageBox.Show("¿Desea Enviar Compras?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No)
			{
				return;
			}
			else
			{
				enviaCostosIndirectos("Compra_CI");			
			}
		}


		private void btnAsiento_CI_Click(object sender, System.EventArgs e)
		{
			if (MessageBox.Show("¿Desea Enviar Asientos?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) 
			{
				return;
			}
			else
			{
				Cursor = Cursors.WaitCursor;
				enviaCostosIndirectos("Asiento_CI");			
				Cursor = Cursors.Default;
			}
		}

		private void ultraButton3_Click(object sender, System.EventArgs e)
		{
			DateTime dtHasta;
			DateTime dtDesde;
			try
			{
				dtDesde = (DateTime) cmbFechaD.Value;
				dtHasta = (DateTime) cmbFechaH.Value;
			}
			catch(Exception ex)
			{
				MessageBox.Show("Error: "+ex);
				return;
			}

			if (MessageBox.Show("¿Desea Habilitar Egresos del lote " + txtparametro.Text + " para reenvio?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			string stExec = string.Format(" update Compra set bloquear = 0 , Mensaje1 = '', bio = '' where RTrim(Mensaje1) = '{0}' "
				+"and fecha >= '{1}' and fecha <= '{2}' and idTipoFactura = {3}",
				txtparametro.Text.Trim(), dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"),8);

			Funcion.EjecutaSQL(cdsAsiento, stExec, true); 
			MessageBox.Show("Egresos del lote " + txtparametro.Text + " habilitados",
				"", MessageBoxButtons.OK,MessageBoxIcon.Question);

		}

		private void label4_Click(object sender, System.EventArgs e)
		{
		
		}

		private void IntCostoIndirecto_Load(object sender, System.EventArgs e)
		{
		
		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			DateTime dtHasta;
			DateTime dtDesde;
			try
			{
				dtDesde = (DateTime) cmbFechaD.Value;
				dtHasta = (DateTime) cmbFechaH.Value;
			}
			catch(Exception ex)
			{
				MessageBox.Show("Error: "+ex);
				return;
			}
			if (MessageBox.Show("¿Desea Habilitar Asiento del lote " + txtparametro.Text + " para reenvio?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			string stExec = string.Format(" update Asiento set Numero = 'Bio' where Numero = '{0}' and fecha >= '{1}' and fecha <= '{2}'",
				txtparametro.Text.Trim(), dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"));

			Funcion.EjecutaSQL(cdsAsiento, stExec, true);
			MessageBox.Show("Asientos del lote " + txtparametro.Text + " habilitados",
				"", MessageBoxButtons.OK,MessageBoxIcon.Question);
		}

		private void ultraButton2_Click(object sender, System.EventArgs e)
		{
			DateTime dtHasta;
			DateTime dtDesde;
			try
			{
				dtDesde = (DateTime) cmbFechaD.Value;
				dtHasta = (DateTime) cmbFechaH.Value;
			}
			catch(Exception ex)
			{
				MessageBox.Show("Error: "+ex);
				return;
			}
			if (MessageBox.Show("¿Desea Habilitar Compras del lote " + txtparametro.Text + " para reenvio?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			string stExec = string.Format(" update Compra set bloquear = 0 , Mensaje1 = '', bio = '' where Mensaje1 = '{0}' and fecha >= '{1}' and fecha <= '{2}' and idTipoFactura = {3}",
				 txtparametro.Text.Trim(), dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"),4);
			Funcion.EjecutaSQL(cdsAsiento, stExec, true);
			MessageBox.Show("Compras del lote " + txtparametro.Text + " habilitados",
				"", MessageBoxButtons.OK,MessageBoxIcon.Question);

		}

		private void ultraButton4_Click(object sender, System.EventArgs e)
		{
			enviaCostosIndirectos("Depreciacion_CI");

		}

		private void ultraButton5_Click(object sender, System.EventArgs e)
		{
			DateTime dtHasta;
			DateTime dtDesde;
			try
			{
				dtDesde = (DateTime) cmbFechaD.Value;
				dtHasta = (DateTime) cmbFechaH.Value;
			}
			catch(Exception ex)
			{
				MessageBox.Show("Error: "+ex);
				return;
			}
			if (MessageBox.Show("¿Desea Habilitar Depreciaciones del lote " + txtparametro.Text + " para reenvio?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			string stExec = string.Format(" update Depreciacion set bio = '' where bio = '{0}' and fecha >= '{1}' and fecha <= '{2}'",
			 txtparametro.Text, dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"));

			Funcion.EjecutaSQL(cdsAsiento, stExec, true);
			MessageBox.Show("Depreciacion del lote " + txtparametro.Text + " habilitados",
				"", MessageBoxButtons.OK,MessageBoxIcon.Question);
		}

		private void btn_ReporteCostosInd_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			DateTime dtDesde = (DateTime) cmbFechaD.Value;
			DateTime dtHasta = (DateTime) cmbFechaH.Value;			
							
			
			string strFiltro = "{TmpCI_Botrosa.FechaDesde} >= #"+ 
				dtDesde.ToString("MM/dd/yyyy") + 
				"# AND {TmpCI_Botrosa.FechaHasta} <= #" + dtHasta.ToString("MM/dd/yyyy") +"#";			

				
			
			Reporte miReporte = new Reporte("CostosIndirectos.rpt",strFiltro);
			miReporte.MdiParent = this.MdiParent;
			miReporte.Show();
			Cursor = Cursors.Default;

		}		
	}
}
