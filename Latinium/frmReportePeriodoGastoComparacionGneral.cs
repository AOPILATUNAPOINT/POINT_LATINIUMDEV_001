using System;
using System.Globalization;
using System.Data;
using System.Data.SqlClient;
using C1.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using System.Web.Mail;
using CrystalDecisions.Shared;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using System.Text;
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;
using System.Security.Cryptography.Xml;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReportePeriodoGastoComparacionGneral.
	/// </summary>
	public class frmReportePeriodoGastoComparacionGneral : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Button btnExcel;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAño;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReportePeriodoGastoComparacionGneral()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReportePeriodoGastoComparacionGneral));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			this.label1 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblContador = new System.Windows.Forms.Label();
			this.btnExcel = new System.Windows.Forms.Button();
			this.txtAño = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnVer = new System.Windows.Forms.Button();
			this.uGridInformacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAño)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 16);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(43, 16);
			this.label1.TabIndex = 1002;
			this.label1.Text = "Periodo";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance1;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(960, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(40, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 999;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.Visible = false;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.lblContador.Location = new System.Drawing.Point(280, 16);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 17);
			this.lblContador.TabIndex = 998;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnExcel
			// 
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(1008, 8);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(75, 24);
			this.btnExcel.TabIndex = 997;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Visible = false;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// txtAño
			// 
			this.txtAño.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAño.Location = new System.Drawing.Point(72, 14);
			this.txtAño.MinValue = 2017;
			this.txtAño.Name = "txtAño";
			this.txtAño.PromptChar = ' ';
			this.txtAño.Size = new System.Drawing.Size(88, 21);
			this.txtAño.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtAño.TabIndex = 995;
			this.txtAño.Click += new System.EventHandler(this.txtAño_Click);
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(176, 13);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(88, 23);
			this.btnVer.TabIndex = 996;
			this.btnVer.Text = "Buscar";
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// uGridInformacion
			// 
			this.uGridInformacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridInformacion.Cursor = System.Windows.Forms.Cursors.Default;
			appearance2.BackColor = System.Drawing.Color.White;
			this.uGridInformacion.DisplayLayout.Appearance = appearance2;
			this.uGridInformacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridInformacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridInformacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.uGridInformacion.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInformacion.DisplayLayout.Override.CellAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 8F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInformacion.DisplayLayout.Override.HeaderAppearance = appearance5;
			this.uGridInformacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.uGridInformacion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInformacion.Location = new System.Drawing.Point(8, 56);
			this.uGridInformacion.Name = "uGridInformacion";
			this.uGridInformacion.Size = new System.Drawing.Size(1072, 376);
			this.uGridInformacion.TabIndex = 1003;
			this.uGridInformacion.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridInformacion_ClickCellButton);
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
			// frmReportePeriodoGastoComparacionGneral
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1088, 446);
			this.Controls.Add(this.uGridInformacion);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.txtAño);
			this.Controls.Add(this.btnVer);
			this.KeyPreview = true;
			this.Name = "frmReportePeriodoGastoComparacionGneral";
			this.Text = "Gasto vs Presupuesto General";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReportePeriodoGastoComparacionGneral_KeyDown);
			this.Load += new System.EventHandler(this.frmReportePeriodoGastoComparacionGneral_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAño)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmReportePeriodoGastoComparacionGneral_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}

		private void frmReportePeriodoGastoComparacionGneral_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "26100102");

			if (!Funcion.Permiso("1767", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			if (miAcceso.BExportar) this.btnExcel.Visible = true;

			this.txtAño.MinValue = 2020;
			this.txtAño.MaxValue = DateTime.Today.Year + 1;
			this.txtAño.Value = DateTime.Today.Year;
			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);
		

		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
	
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			
			this.uGridInformacion.DataSource = null;
			string sSQL = string.Format("Exec [ReportePresupuestoVsRealGeneralTOTAL] {0}, {1}", (int)this.txtAño.Value, 0);
			this.uGridInformacion.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);

			this.DiseñoGridTurno(this.uGridInformacion, true, false, "");
			this.DiseñoGridColor(this.uGridInformacion);

			this.lblContador.Text = this.uGridInformacion.Rows.Count +" - Registros Encontrados";
		}

		private void txtAño_Click(object sender, System.EventArgs e)
		{
			this.txtAño.SelectAll();
		}
		private void DiseñoGridColor(Infragistics.Win.UltraWinGrid.UltraGrid uGrid)
		{

			string sColumna = "";
			int iContColumnas = (int)this.uGridInformacion.DisplayLayout.Bands[0].Columns.Count;	
		
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInformacion.Rows.All)
			{
				for (int i = 2; i < iContColumnas; i++)
				{

					sColumna = this.uGridInformacion.DisplayLayout.Bands[0].Columns[i].Header.Caption;			

					//	if (dr.Cells["Enero"].Value.ToString().Trim().Length == 0)  dr.Cells["0-0"].Value = 0;
					//					if (dr.Cells["1-5"].Value.ToString().Trim().Length == 0) dr.Cells["1-5"].Value = 0;
					//					if (dr.Cells["6-15"].Value.ToString().Trim().Length == 0) dr.Cells["6-15"].Value = 0;
					//					if (dr.Cells["16-30"].Value.ToString().Trim().Length == 0)  dr.Cells["16-30"].Value = 0;
					//
					//					if (dr.Cells["31-60"].Value.ToString().Trim().Length == 0)  dr.Cells["31-60"].Value = 0;
					//					if (dr.Cells["61-90"].Value.ToString().Trim().Length == 0)  dr.Cells["61-90"].Value = 0;
					//					if (dr.Cells["91-99"].Value.ToString().Trim().Length == 0)  dr.Cells["91-99"].Value = 0;
					//					if (dr.Cells["100-180"].Value.ToString().Trim().Length == 0)  dr.Cells["100-180"].Value = 0;
					//
					//					double dSumC=  Convert.ToDouble(dr.Cells["0-0"].Value) +  Convert.ToDouble(dr.Cells["1-5"].Value)  + 
					//						Convert.ToDouble(dr.Cells["6-15"].Value) +Convert.ToDouble(dr.Cells["16-30"].Value) ;
					#region Enero
					if( sColumna == "Enero" || sColumna == "Gasto_Enero" || sColumna == "Total_Enero" )
					{
						decimal sPorce = 0;
						
						if ((decimal) dr.Cells["Enero"].Value> 0)
						{
							sPorce = ((decimal) dr.Cells["Gasto_Enero"].Value * 100 )/(decimal) dr.Cells["Enero"].Value;

						}
						else
						{
							sPorce = 0;
						}
						
						//							double dNumero = Convert.ToDouble ( dr.Cells[sColumna].Value.ToString ());
						decimal dNumero = sPorce;

						
								
						if(  Math.Round(dNumero, 2) < 50)
						{
					
							dr.Cells[sColumna].Appearance.BackColor = Color.Green;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.Green;
						}
						if(Math.Round(dNumero, 2) >=50 &&  Math.Round(dNumero, 2) <= 80)
						{
					
							dr.Cells[sColumna].Appearance.BackColor = Color.Orange;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.Orange;
						}
						if(Math.Round(dNumero, 2) >80 )
						{
					
							dr.Cells[sColumna].Appearance.BackColor = Color.Red;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.Red;
						}
					}
					#endregion Rango 0-0

					#region Febrero
					if( sColumna == "Febrero" || sColumna == "Gasto_Febrero" || sColumna == "Total_Febrero" )
					{
						decimal sPorce = 0;
						
						if ((decimal) dr.Cells["Febrero"].Value> 0)
						{
							sPorce = ((decimal) dr.Cells["Gasto_Febrero"].Value * 100 )/(decimal) dr.Cells["Febrero"].Value;

						}
						else
						{
							sPorce = 0;
						}
						
						//							double dNumero = Convert.ToDouble ( dr.Cells[sColumna].Value.ToString ());
						decimal dNumero = sPorce;
								
						if(  Math.Round(dNumero, 2) < 50)
						{
					
							dr.Cells[sColumna].Appearance.BackColor = Color.Green;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.Green;
						}
						if(Math.Round(dNumero, 2) >=50 &&  Math.Round(dNumero, 2) <= 80)
						{
					
							dr.Cells[sColumna].Appearance.BackColor = Color.Orange;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.Orange;
						}
						if(Math.Round(dNumero, 2) >80 )
						{
					
							dr.Cells[sColumna].Appearance.BackColor = Color.Red;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.Red;
						}
					}
					#endregion Febrero

					#region Marzo
					if( sColumna == "Marzo" || sColumna == "Gasto_Marzo" || sColumna == "Total_Marzo" )
					{
						decimal sPorce = 0;
						
						if ((decimal) dr.Cells["Marzo"].Value> 0)
						{
							sPorce = ((decimal) dr.Cells["Gasto_Marzo"].Value * 100 )/(decimal) dr.Cells["Marzo"].Value;

						}
						else
						{
							sPorce = 0;
						}
						
						//							double dNumero = Convert.ToDouble ( dr.Cells[sColumna].Value.ToString ());
						decimal dNumero = sPorce;
								
						if(  Math.Round(dNumero, 2) < 50)
						{
					
							dr.Cells[sColumna].Appearance.BackColor = Color.Green;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.Green;
						}
						if(Math.Round(dNumero, 2) >=50 &&  Math.Round(dNumero, 2) <= 80)
						{
					
							dr.Cells[sColumna].Appearance.BackColor = Color.Orange;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.Orange;
						}
						if(Math.Round(dNumero, 2) >80 )
						{
					
							dr.Cells[sColumna].Appearance.BackColor = Color.Red;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.Red;
						}
					}
					#endregion Febrero

					#region Abril
					if( sColumna == "Abril" || sColumna == "Gasto_Abril" || sColumna == "Total_Abril" )
					{
						decimal sPorce = 0;
						
						if ((decimal) dr.Cells["Abril"].Value> 0)
						{
							sPorce = ((decimal) dr.Cells["Gasto_Abril"].Value * 100 )/(decimal) dr.Cells["Abril"].Value;

						}
						else
						{
							sPorce = 0;
						}
						
						//							double dNumero = Convert.ToDouble ( dr.Cells[sColumna].Value.ToString ());
						decimal dNumero = sPorce;
								
						if(  Math.Round(dNumero, 2) < 50)
						{
					
							dr.Cells[sColumna].Appearance.BackColor = Color.Green;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.Green;
						}
						if(Math.Round(dNumero, 2) >=50 &&  Math.Round(dNumero, 2) <= 80)
						{
					
							dr.Cells[sColumna].Appearance.BackColor = Color.Orange;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.Orange;
						}
						if(Math.Round(dNumero, 2) >80 )
						{
					
							dr.Cells[sColumna].Appearance.BackColor = Color.Red;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.Red;
						}
					}
					#endregion Abril

					#region Mayo
					if( sColumna == "Mayo" || sColumna == "Gasto_Mayo" || sColumna == "Total_Mayo" )
					{
						decimal sPorce = 0;
						
						if ((decimal) dr.Cells["Mayo"].Value> 0)
						{
							sPorce = ((decimal) dr.Cells["Gasto_Mayo"].Value * 100 )/(decimal) dr.Cells["Mayo"].Value;

						}
						else
						{
							sPorce = 0;
						}
						
						//							double dNumero = Convert.ToDouble ( dr.Cells[sColumna].Value.ToString ());
						decimal dNumero = sPorce;
								
						if(  Math.Round(dNumero, 2) < 50)
						{
					
							dr.Cells[sColumna].Appearance.BackColor = Color.Green;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.Green;
						}
						if(Math.Round(dNumero, 2) >=50 &&  Math.Round(dNumero, 2) <= 80)
						{
					
							dr.Cells[sColumna].Appearance.BackColor = Color.Orange;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.Orange;
						}
						if(Math.Round(dNumero, 2) >80 )
						{
					
							dr.Cells[sColumna].Appearance.BackColor = Color.Red;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.Red;
						}
					}
					#endregion Abril

					#region Junio
					if( sColumna == "Junio" || sColumna == "Gasto_Junio" || sColumna == "Total_Junio" )
					{
						decimal sPorce = 0;
						
						if ((decimal) dr.Cells["Junio"].Value> 0)
						{
							sPorce = ((decimal) dr.Cells["Gasto_Junio"].Value * 100 )/(decimal) dr.Cells["Junio"].Value;

						}
						else
						{
							sPorce = 0;
						}
						
						//							double dNumero = Convert.ToDouble ( dr.Cells[sColumna].Value.ToString ());
						decimal dNumero = sPorce;
								
						if(  Math.Round(dNumero, 2) < 50)
						{
					
							dr.Cells[sColumna].Appearance.BackColor = Color.Green;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.Green;
						}
						if(Math.Round(dNumero, 2) >=50 &&  Math.Round(dNumero, 2) <= 80)
						{
					
							dr.Cells[sColumna].Appearance.BackColor = Color.Orange;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.Orange;
						}
						if(Math.Round(dNumero, 2) >80 )
						{
					
							dr.Cells[sColumna].Appearance.BackColor = Color.Red;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.Red;
						}
					}
					#endregion Abril

					#region Julio
					if( sColumna == "Julio" || sColumna == "Gasto_Julio" || sColumna == "Total_Julio" )
					{
						decimal sPorce = 0;
						
						if ((decimal) dr.Cells["Julio"].Value> 0)
						{
							sPorce = ((decimal) dr.Cells["Gasto_Julio"].Value * 100 )/(decimal) dr.Cells["Julio"].Value;

						}
						else
						{
							sPorce = 0;
						}
						
						//							double dNumero = Convert.ToDouble ( dr.Cells[sColumna].Value.ToString ());
						decimal dNumero = sPorce;
								
						if(  Math.Round(dNumero, 2) < 50)
						{
					
							dr.Cells[sColumna].Appearance.BackColor = Color.Green;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.Green;
						}
						if(Math.Round(dNumero, 2) >=50 &&  Math.Round(dNumero, 2) <= 80)
						{
					
							dr.Cells[sColumna].Appearance.BackColor = Color.Orange;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.Orange;
						}
						if(Math.Round(dNumero, 2) >80 )
						{
					
							dr.Cells[sColumna].Appearance.BackColor = Color.Red;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.Red;
						}
					}
					#endregion Julio

					#region Agosto
					if( sColumna == "Agosto" || sColumna == "Gasto_Agosto" || sColumna == "Total_Agosto" )
					{
						decimal sPorce = 0;
						
						if ((decimal) dr.Cells["Agosto"].Value> 0)
						{
							sPorce = ((decimal) dr.Cells["Gasto_Agosto"].Value * 100 )/(decimal) dr.Cells["Agosto"].Value;

						}
						else
						{
							sPorce = 0;
						}
						
						//							double dNumero = Convert.ToDouble ( dr.Cells[sColumna].Value.ToString ());
						decimal dNumero = sPorce;
								
						if(  Math.Round(dNumero, 2) < 50)
						{
					
							dr.Cells[sColumna].Appearance.BackColor = Color.Green;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.Green;
						}
						if(Math.Round(dNumero, 2) >=50 &&  Math.Round(dNumero, 2) <= 80)
						{
					
							dr.Cells[sColumna].Appearance.BackColor = Color.Orange;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.Orange;
						}
						if(Math.Round(dNumero, 2) >80 )
						{
					
							dr.Cells[sColumna].Appearance.BackColor = Color.Red;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.Red;
						}
					}
					#endregion Agosto

					#region Septiembre
					if( sColumna == "Septiembre" || sColumna == "Gasto_Septiembre" || sColumna == "Total_Septiembre" )
					{
						decimal sPorce = 0;
						
						if ((decimal) dr.Cells["Septiembre"].Value> 0)
						{
							sPorce = ((decimal) dr.Cells["Gasto_Septiembre"].Value * 100 )/(decimal) dr.Cells["Septiembre"].Value;

						}
						else
						{
							sPorce = 0;
						}
						
						//							double dNumero = Convert.ToDouble ( dr.Cells[sColumna].Value.ToString ());
						decimal dNumero = sPorce;
								
						if(  Math.Round(dNumero, 2) < 50)
						{
					
							dr.Cells[sColumna].Appearance.BackColor = Color.Green;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.Green;
						}
						if(Math.Round(dNumero, 2) >=50 &&  Math.Round(dNumero, 2) <= 80)
						{
					
							dr.Cells[sColumna].Appearance.BackColor = Color.Orange;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.Orange;
						}
						if(Math.Round(dNumero, 2) >80 )
						{
					
							dr.Cells[sColumna].Appearance.BackColor = Color.Red;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.Red;
						}
					}
					#endregion Septiembre

					#region Octubre
					if( sColumna == "Octubre" || sColumna == "Gasto_Octubre" || sColumna == "Total_Octubre" )
					{
						decimal sPorce = 0;
						
						if ((decimal) dr.Cells["Octubre"].Value> 0)
						{
							sPorce = ((decimal) dr.Cells["Gasto_Octubre"].Value * 100 )/(decimal) dr.Cells["Octubre"].Value;

						}
						else
						{
							sPorce = 0;
						}
						
						//							double dNumero = Convert.ToDouble ( dr.Cells[sColumna].Value.ToString ());
						decimal dNumero = sPorce;
								
						if(  Math.Round(dNumero, 2) < 50)
						{
					
							dr.Cells[sColumna].Appearance.BackColor = Color.Green;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.Green;
						}
						if(Math.Round(dNumero, 2) >=50 &&  Math.Round(dNumero, 2) <= 80)
						{
					
							dr.Cells[sColumna].Appearance.BackColor = Color.Orange;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.Orange;
						}
						if(Math.Round(dNumero, 2) >80 )
						{
					
							dr.Cells[sColumna].Appearance.BackColor = Color.Red;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.Red;
						}
					}
					#endregion Octubre

					#region Noviembre
					if( sColumna == "Noviembre" || sColumna == "Gasto_Noviembre" || sColumna == "Total_Noviembre" )
					{
						decimal sPorce = 0;
						
						if ((decimal) dr.Cells["Noviembre"].Value> 0)
						{
							sPorce = ((decimal) dr.Cells["Gasto_Noviembre"].Value * 100 )/(decimal) dr.Cells["Noviembre"].Value;

						}
						else
						{
							sPorce = 0;
						}
						
						//							double dNumero = Convert.ToDouble ( dr.Cells[sColumna].Value.ToString ());
						decimal dNumero = sPorce;
								
						if(  Math.Round(dNumero, 2) < 50)
						{
					
							dr.Cells[sColumna].Appearance.BackColor = Color.Green;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.Green;
						}
						if(Math.Round(dNumero, 2) >=50 &&  Math.Round(dNumero, 2) <= 80)
						{
					
							dr.Cells[sColumna].Appearance.BackColor = Color.Orange;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.Orange;
						}
						if(Math.Round(dNumero, 2) >80 )
						{
					
							dr.Cells[sColumna].Appearance.BackColor = Color.Red;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.Red;
						}
					}
					#endregion Noviembre

					#region Diciembre
					if( sColumna == "Diciembre" || sColumna == "Gasto_Diciembre" || sColumna == "Total_Diciembre" )
					{
						decimal sPorce = 0;
						
						if ((decimal) dr.Cells["Diciembre"].Value> 0)
						{
							sPorce = ((decimal) dr.Cells["Gasto_Diciembre"].Value * 100 )/(decimal) dr.Cells["Diciembre"].Value;

						}
						else
						{
							sPorce = 0;
						}
						
						//							double dNumero = Convert.ToDouble ( dr.Cells[sColumna].Value.ToString ());
						decimal dNumero = sPorce;
								
						if(  Math.Round(dNumero, 2) < 50)
						{
					
							dr.Cells[sColumna].Appearance.BackColor = Color.Green;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.Green;
						}
						if(Math.Round(dNumero, 2) >=50 &&  Math.Round(dNumero, 2) <= 80)
						{
					
							dr.Cells[sColumna].Appearance.BackColor = Color.Orange;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.Orange;
						}
						if(Math.Round(dNumero, 2) >80 )
						{
					
							dr.Cells[sColumna].Appearance.BackColor = Color.Red;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.Red;
						}
					}
					#endregion Diciembre

					
				}
			}
		}


		private void DiseñoGridTurno(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, bool bIniciaMes, bool bSumatorias, string sColumna)
		{
			try
			{				
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["idCuenta"].Hidden = true;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Codigo"].Hidden = false;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Descripcion"].Hidden = false;
				//				
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Codigo"].Width = 150;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Descripcion"].Width = 250;

				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Codigo"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Descripcion"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
				
				int iContCerrado = 0;
				int iContAbierto = 0;

				if (this.uGridInformacion.Rows.Count > 0)
				{				
					int iContColumnas = (int)this.uGridInformacion.DisplayLayout.Bands[0].Columns.Count;					
					
					string nombrCol = "";
					int j = 4;

					for ( j = j  ; j < iContColumnas; j=j+3)
					{
						nombrCol = this.uGridInformacion.DisplayLayout.Bands[0].Columns[j].Header.Caption;
						this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].Style = ColumnStyle.Button;
						this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].Width = 100;
						this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
						this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].Format = "#,##0.00";
						this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].Header.Caption = nombrCol;		
					}
				}
				#region Diseño Grid
				if (uGrid.Rows.Count > 0)
				{
					int iColIni = 3;

				
					if (uGrid.Rows.Count > 0)
					{
						string nombrCol = "";
							
						int iColumnas = (int)uGrid.DisplayLayout.Bands[0].Columns.Count;
				
						for (int j = iColIni; j < iColumnas; j++)
						{
							
							nombrCol = uGrid.DisplayLayout.Bands[0].Columns[j].Header.Caption;
										
							uGrid.DisplayLayout.Bands[0].Columns[nombrCol].CellActivation = Activation.ActivateOnly;
							uGrid.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
							
							uGrid.DisplayLayout.Bands[0].Columns[nombrCol].Width = 80;
							uGrid.DisplayLayout.Bands[0].Summaries.Add(nombrCol, Infragistics.Win.UltraWinGrid.SummaryType.Sum, uGrid.DisplayLayout.Bands[0].Columns[nombrCol], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
							//							int a = Infragistics.Win.UltraWinGrid.SummaryType.Sum;			
							uGrid.DisplayLayout.Bands[0].Summaries[nombrCol].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
							uGrid.DisplayLayout.Bands[0].Summaries[nombrCol].DisplayFormat = "{0: #,##0.00}";
							
						}					
					}				
				}	
				#endregion Diseño Grid
				//				this.lblContador.Text = iContAbierto.ToString()+ " Cajas Abiertas - " + iContCerrado.ToString() + " Cajas Cerradas " ;
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void uGridInformacion_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			string sColumna = e.Cell.Column.Key;
			string sFecha ="";

			if (sColumna == "Gasto_Enero"	) sFecha ="01/01/" + Convert.ToString( this.txtAño.Value);
			if (sColumna == "Gasto_Febrero"	) sFecha ="01/02/" + Convert.ToString( this.txtAño.Value);
			if (sColumna == "Gasto_Marzo"	) sFecha ="01/03/" + Convert.ToString( this.txtAño.Value);
			if (sColumna == "Gasto_Abril"	) sFecha ="01/04/" + Convert.ToString( this.txtAño.Value);
			if (sColumna == "Gasto_Mayo") sFecha ="01/05/" + Convert.ToString( this.txtAño.Value);
			if (sColumna == "Gasto_Junio"	) sFecha ="01/06/" + Convert.ToString( this.txtAño.Value);
			if (sColumna == "Gasto_Julio"	) sFecha ="01/07/" + Convert.ToString( this.txtAño.Value);
			if (sColumna == "Gasto_Agosto") sFecha ="01/08/" + Convert.ToString( this.txtAño.Value);	
			if (sColumna == "Gasto_Septiembre") sFecha ="01/09/" + Convert.ToString( this.txtAño.Value);
			if (sColumna == "Gasto_Octubre"	) sFecha ="01/10/" + Convert.ToString( this.txtAño.Value);
			if (sColumna == "Gasto_Noviembre"	) sFecha ="01/11/" + Convert.ToString( this.txtAño.Value);
			if (sColumna == "Gasto_Diciembre"	) sFecha ="01/12/" + Convert.ToString( this.txtAño.Value);

			if (e.Cell.Row.Cells[sColumna].Value == System.DBNull.Value)
			{
				MessageBox.Show(string.Format("No hay  registrado en la Columna : {0}", sColumna), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			else
			{
				string sMes =sColumna;
				sMes.Replace("Gasto_", "");
				using (frmReportePeriodoGastoDetalleComparacion CCJ = new frmReportePeriodoGastoDetalleComparacion ( (int)e.Cell.Row.Cells["idCuenta"].Value, 
								 sFecha, miAcceso.BExportar, sMes, (string)e.Cell.Row.Cells["Codigo"].Value , (string)e.Cell.Row.Cells["Descripcion"].Value, 0))
				{
					if (DialogResult.OK == CCJ.ShowDialog())
					{
						
					}
				}
			}
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridInformacion);
		}

	}
}
