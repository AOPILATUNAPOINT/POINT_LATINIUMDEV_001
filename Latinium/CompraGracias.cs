using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Globalization;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de CompraGracias.
	/// </summary>
	public class CompraGracias : DevExpress.XtraEditors.XtraForm
	{
		private System.ComponentModel.IContainer components;

		private double dTotal = 0;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtVuelto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtRecibido;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor Vuelto1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor Vuelto5;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor Vuelto10;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor Recibido1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor Recibido5;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor Recibido10;
		private Infragistics.Win.Misc.UltraLabel lblVuelto;
		private Infragistics.Win.Misc.UltraLabel lblRecibido;
		private System.Windows.Forms.Timer timer1;
		private AxMSCommLib.AxMSComm comDisplay;
		private AxMSCommLib.AxMSComm comCaja;
		private C1.Data.C1DataSet cdsSeteo;
		private double dRecibido = 0;
		public CompraGracias(int ITiempo, double DTotal, double DRecibido)
		{
			InitializeComponent();
			dTotal = DTotal;
			dRecibido = DRecibido;
			timer1.Interval = ITiempo * 1000;
			if (!timer1.Enabled)
				timer1.Enabled=true;    // Activamos el Timer si no esta Enabled (Activado)
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
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(CompraGracias));
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtVuelto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.lblVuelto = new Infragistics.Win.Misc.UltraLabel();
			this.lblRecibido = new Infragistics.Win.Misc.UltraLabel();
			this.txtRecibido = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.Vuelto1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.Vuelto5 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.Vuelto10 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.Recibido1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.Recibido5 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.Recibido10 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.timer1 = new System.Windows.Forms.Timer(this.components);
			this.comDisplay = new AxMSCommLib.AxMSComm();
			this.comCaja = new AxMSCommLib.AxMSComm();
			this.cdsSeteo = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVuelto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRecibido)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.Vuelto1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.Vuelto5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.Vuelto10)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.Recibido1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.Recibido5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.Recibido10)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.comDisplay)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.comCaja)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			this.SuspendLayout();
			// 
			// txtTotal
			// 
			appearance1.BackColorDisabled = System.Drawing.Color.White;
			appearance1.ForeColorDisabled = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(192)));
			this.txtTotal.Appearance = appearance1;
			this.txtTotal.BorderStyle = Infragistics.Win.UIElementBorderStyle.None;
			this.txtTotal.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Button3D;
			this.txtTotal.Enabled = false;
			this.txtTotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 48F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(256, 24);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(368, 84);
			this.txtTotal.TabIndex = 0;
			this.txtTotal.ValueChanged += new System.EventHandler(this.txtTotal_ValueChanged);
			// 
			// txtVuelto
			// 
			appearance2.BackColorDisabled = System.Drawing.Color.White;
			appearance2.ForeColorDisabled = System.Drawing.Color.Green;
			this.txtVuelto.Appearance = appearance2;
			this.txtVuelto.BorderStyle = Infragistics.Win.UIElementBorderStyle.None;
			this.txtVuelto.Enabled = false;
			this.txtVuelto.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtVuelto.FormatString = "#,##0.00";
			this.txtVuelto.Location = new System.Drawing.Point(288, 144);
			this.txtVuelto.Name = "txtVuelto";
			this.txtVuelto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtVuelto.PromptChar = ' ';
			this.txtVuelto.Size = new System.Drawing.Size(336, 64);
			this.txtVuelto.TabIndex = 1;
			// 
			// ultraLabel1
			// 
			appearance3.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(192)));
			appearance3.ForeColorDisabled = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(192)));
			this.ultraLabel1.Appearance = appearance3;
			this.ultraLabel1.Font = new System.Drawing.Font("Microsoft Sans Serif", 48F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraLabel1.Location = new System.Drawing.Point(8, 24);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(248, 80);
			this.ultraLabel1.TabIndex = 2;
			this.ultraLabel1.Text = "TOTAL";
			// 
			// lblVuelto
			// 
			appearance4.ForeColor = System.Drawing.Color.Green;
			this.lblVuelto.Appearance = appearance4;
			this.lblVuelto.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblVuelto.Location = new System.Drawing.Point(16, 144);
			this.lblVuelto.Name = "lblVuelto";
			this.lblVuelto.Size = new System.Drawing.Size(224, 64);
			this.lblVuelto.TabIndex = 3;
			this.lblVuelto.Text = "VUELTO";
			// 
			// lblRecibido
			// 
			appearance5.ForeColor = System.Drawing.Color.Maroon;
			this.lblRecibido.Appearance = appearance5;
			this.lblRecibido.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblRecibido.Location = new System.Drawing.Point(16, 232);
			this.lblRecibido.Name = "lblRecibido";
			this.lblRecibido.Size = new System.Drawing.Size(264, 64);
			this.lblRecibido.TabIndex = 5;
			this.lblRecibido.Text = "RECIBIDO";
			// 
			// txtRecibido
			// 
			appearance6.BackColorDisabled = System.Drawing.Color.White;
			appearance6.ForeColorDisabled = System.Drawing.Color.Maroon;
			this.txtRecibido.Appearance = appearance6;
			this.txtRecibido.BorderStyle = Infragistics.Win.UIElementBorderStyle.None;
			this.txtRecibido.Enabled = false;
			this.txtRecibido.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtRecibido.FormatString = "#,##0.00";
			this.txtRecibido.Location = new System.Drawing.Point(288, 232);
			this.txtRecibido.Name = "txtRecibido";
			this.txtRecibido.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtRecibido.PromptChar = ' ';
			this.txtRecibido.Size = new System.Drawing.Size(336, 64);
			this.txtRecibido.TabIndex = 4;
			// 
			// Vuelto1
			// 
			appearance7.BackColorDisabled = System.Drawing.Color.White;
			appearance7.ForeColorDisabled = System.Drawing.Color.Green;
			this.Vuelto1.Appearance = appearance7;
			this.Vuelto1.BorderStyle = Infragistics.Win.UIElementBorderStyle.None;
			this.Vuelto1.Enabled = false;
			this.Vuelto1.Font = new System.Drawing.Font("Microsoft Sans Serif", 27.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.Vuelto1.FormatString = "#,##0.00";
			this.Vuelto1.Location = new System.Drawing.Point(72, 152);
			this.Vuelto1.Name = "Vuelto1";
			this.Vuelto1.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.Vuelto1.PromptChar = ' ';
			this.Vuelto1.Size = new System.Drawing.Size(176, 50);
			this.Vuelto1.TabIndex = 6;
			this.Vuelto1.Visible = false;
			// 
			// Vuelto5
			// 
			appearance8.BackColorDisabled = System.Drawing.Color.White;
			appearance8.ForeColorDisabled = System.Drawing.Color.Green;
			this.Vuelto5.Appearance = appearance8;
			this.Vuelto5.BorderStyle = Infragistics.Win.UIElementBorderStyle.None;
			this.Vuelto5.Enabled = false;
			this.Vuelto5.Font = new System.Drawing.Font("Microsoft Sans Serif", 27.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.Vuelto5.FormatString = "#,##0.00";
			this.Vuelto5.Location = new System.Drawing.Point(248, 152);
			this.Vuelto5.Name = "Vuelto5";
			this.Vuelto5.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.Vuelto5.PromptChar = ' ';
			this.Vuelto5.Size = new System.Drawing.Size(176, 50);
			this.Vuelto5.TabIndex = 7;
			this.Vuelto5.Visible = false;
			// 
			// Vuelto10
			// 
			appearance9.BackColorDisabled = System.Drawing.Color.White;
			appearance9.ForeColorDisabled = System.Drawing.Color.Green;
			this.Vuelto10.Appearance = appearance9;
			this.Vuelto10.BorderStyle = Infragistics.Win.UIElementBorderStyle.None;
			this.Vuelto10.Enabled = false;
			this.Vuelto10.Font = new System.Drawing.Font("Microsoft Sans Serif", 27.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.Vuelto10.FormatString = "#,##0.00";
			this.Vuelto10.Location = new System.Drawing.Point(424, 152);
			this.Vuelto10.Name = "Vuelto10";
			this.Vuelto10.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.Vuelto10.PromptChar = ' ';
			this.Vuelto10.Size = new System.Drawing.Size(200, 50);
			this.Vuelto10.TabIndex = 8;
			this.Vuelto10.Visible = false;
			// 
			// Recibido1
			// 
			appearance10.BackColorDisabled = System.Drawing.Color.White;
			appearance10.ForeColorDisabled = System.Drawing.Color.Maroon;
			this.Recibido1.Appearance = appearance10;
			this.Recibido1.BorderStyle = Infragistics.Win.UIElementBorderStyle.None;
			this.Recibido1.Enabled = false;
			this.Recibido1.Font = new System.Drawing.Font("Microsoft Sans Serif", 27.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.Recibido1.FormatString = "#,##0.00";
			this.Recibido1.Location = new System.Drawing.Point(72, 240);
			this.Recibido1.Name = "Recibido1";
			this.Recibido1.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.Recibido1.PromptChar = ' ';
			this.Recibido1.Size = new System.Drawing.Size(176, 50);
			this.Recibido1.TabIndex = 9;
			this.Recibido1.Visible = false;
			// 
			// Recibido5
			// 
			appearance11.BackColorDisabled = System.Drawing.Color.White;
			appearance11.ForeColorDisabled = System.Drawing.Color.Maroon;
			this.Recibido5.Appearance = appearance11;
			this.Recibido5.BorderStyle = Infragistics.Win.UIElementBorderStyle.None;
			this.Recibido5.Enabled = false;
			this.Recibido5.Font = new System.Drawing.Font("Microsoft Sans Serif", 27.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.Recibido5.FormatString = "#,##0.00";
			this.Recibido5.Location = new System.Drawing.Point(248, 240);
			this.Recibido5.Name = "Recibido5";
			this.Recibido5.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.Recibido5.PromptChar = ' ';
			this.Recibido5.Size = new System.Drawing.Size(168, 50);
			this.Recibido5.TabIndex = 10;
			this.Recibido5.Visible = false;
			// 
			// Recibido10
			// 
			appearance12.BackColorDisabled = System.Drawing.Color.White;
			appearance12.ForeColorDisabled = System.Drawing.Color.Maroon;
			this.Recibido10.Appearance = appearance12;
			this.Recibido10.BorderStyle = Infragistics.Win.UIElementBorderStyle.None;
			this.Recibido10.Enabled = false;
			this.Recibido10.Font = new System.Drawing.Font("Microsoft Sans Serif", 27.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.Recibido10.FormatString = "#,##0.00";
			this.Recibido10.Location = new System.Drawing.Point(424, 240);
			this.Recibido10.Name = "Recibido10";
			this.Recibido10.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.Recibido10.PromptChar = ' ';
			this.Recibido10.Size = new System.Drawing.Size(200, 50);
			this.Recibido10.TabIndex = 11;
			this.Recibido10.Visible = false;
			// 
			// timer1
			// 
			this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
			// 
			// comDisplay
			// 
			this.comDisplay.Enabled = true;
			this.comDisplay.Location = new System.Drawing.Point(16, 104);
			this.comDisplay.Name = "comDisplay";
			this.comDisplay.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("comDisplay.OcxState")));
			this.comDisplay.Size = new System.Drawing.Size(38, 38);
			this.comDisplay.TabIndex = 12;
			this.comDisplay.OnComm += new System.EventHandler(this.axMSComm1_OnComm);
			// 
			// comCaja
			// 
			this.comCaja.Enabled = true;
			this.comCaja.Location = new System.Drawing.Point(64, 104);
			this.comCaja.Name = "comCaja";
			this.comCaja.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("comCaja.OcxState")));
			this.comCaja.Size = new System.Drawing.Size(38, 38);
			this.comCaja.TabIndex = 13;
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
			// 
			// CompraGracias
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.BackColor = System.Drawing.Color.White;
			this.ClientSize = new System.Drawing.Size(640, 320);
			this.ControlBox = false;
			this.Controls.Add(this.comCaja);
			this.Controls.Add(this.comDisplay);
			this.Controls.Add(this.Recibido10);
			this.Controls.Add(this.Recibido5);
			this.Controls.Add(this.Recibido1);
			this.Controls.Add(this.Vuelto10);
			this.Controls.Add(this.Vuelto5);
			this.Controls.Add(this.Vuelto1);
			this.Controls.Add(this.lblRecibido);
			this.Controls.Add(this.txtRecibido);
			this.Controls.Add(this.lblVuelto);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.txtVuelto);
			this.Controls.Add(this.txtTotal);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "CompraGracias";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Load += new System.EventHandler(this.CompraGracias_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVuelto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRecibido)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.Vuelto1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.Vuelto5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.Vuelto10)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.Recibido1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.Recibido5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.Recibido10)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.comDisplay)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.comCaja)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		CultureInfo us = new CultureInfo("en-US");
		private void DisplayMostrar(double dVuelto, double dTotal)
		{
			#region Display en puerto serie
			bool bDisplay = Funcion.bEjecutaSQL(cdsSeteo, "Exec SeteoGExiste 'HSD'");
			if (bDisplay)
			{
				try
				{
					char cEsc = Convert.ToChar(27);
					char cEnter = Convert.ToChar(13);
					string stLinea1 = cEsc + "QA T " + dTotal.ToString("0.00") 
						+ " V " + dVuelto.ToString("0.00");
					string stLinea2 = cEsc + "QBDESARROLLO INFOELECT";
					int iPuertoD = Funcion.iEscalarSQL(cdsSeteo, "Select cast(Valor as int) From SeteoG Where Codigo = 'HSD'");
					comDisplay.CommPort = short.Parse(iPuertoD.ToString());
					comDisplay.PortOpen = true;
					comDisplay.Output = stLinea1 + cEnter;
					comDisplay.Output = stLinea2 + cEnter;
				}
				catch(System.Exception ex)
				{
					MessageBox.Show(ex.Message, "Error en puerto de Display");
					return;
				}
			}
			#endregion Display en puerto serie
		}

		private void CompraGracias_Load(object sender, System.EventArgs e)
		{
			double dVuelto = dRecibido - dTotal;

			#region Caja de dinero en puerto serie
			bool bCajaDinero = Funcion.bEjecutaSQL(cdsSeteo, "Exec SeteoGExiste 'HSC'");
			if (bCajaDinero)
			{
				try
				{
					int iPuertoC = Funcion.iEscalarSQL(cdsSeteo, "Select cast(Valor as int) From SeteoG Where Codigo = 'HSC'");
					comCaja.CommPort = short.Parse(iPuertoC.ToString());
					comCaja.PortOpen = true;
					comCaja.Output = "0000000000";
				}
				catch(System.Exception ex)
				{
					MessageBox.Show(ex.Message, "Error en puerto de Caja de Dinero");
					return;
				}
			}
			#endregion Caja de dinero en puerto serie
			
			this.txtTotal.Value = dTotal;
			if (dRecibido < dTotal)
			{
				Vuelto1.Visible = true;
				Vuelto5.Visible = true;
				Vuelto10.Visible = true;
				Recibido1.Visible = true;
				Recibido5.Visible = true;
				Recibido10.Visible = true;

				lblRecibido.Text = "R";
				lblVuelto.Text = "V";
				dRecibido = Math.Floor(dTotal) + 1;
				dVuelto = dRecibido - dTotal;
				Vuelto1.Value = Math.Round(dVuelto, 2);
				Recibido1.Value = Math.Round(dRecibido, 2);

				double dCinco = ((int) ((dTotal)/5)) * 5.0 + 5.0;
				dRecibido = dCinco;
				dVuelto = dRecibido - dTotal;
				Vuelto5.Value = Math.Round(dVuelto, 2);
				Recibido5.Value = Math.Round(dRecibido, 2);
				
				dRecibido = dCinco + 5;
				dVuelto = dRecibido - dTotal;
				Vuelto10.Value = Math.Round(dVuelto, 2);
				Recibido10.Value = Math.Round(dRecibido, 2);
				
				txtVuelto.Visible = false;
				txtRecibido.Visible = false;

				DisplayMostrar(dVuelto, dTotal);
				return;
			}
			this.txtVuelto.Value = Math.Round(dVuelto, 2);
			this.txtRecibido.Value = Math.Round(dRecibido, 2);
			DisplayMostrar(dVuelto, dTotal);
		}

		private void txtTotal_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void timer1_Tick(object sender, System.EventArgs e)
		{
			timer1.Stop();     // Se para el timer.
			this.Close();      // Cerramos el formulario.
		}

		private void axMSComm1_OnComm(object sender, System.EventArgs e)
		{
		}
	}
}
