using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Text;
using C1.Data;
using System.Globalization;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCre_CarteraRefinanciadaPichincha.
	/// </summary>
	public class frmCre_CarteraRefinanciadaPichincha : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private System.Windows.Forms.Button btnVer;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnGenerar;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridOperaciones;
		private System.Windows.Forms.Button btnSeleccionar;
		private System.Windows.Forms.Button btnActualizar;
		private System.Windows.Forms.Label lblContador;
		private Acceso miAcceso;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCre_CarteraRefinanciadaPichincha()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCre_CarteraRefinanciadaPichincha));
			this.uGridOperaciones = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.lblBodega = new System.Windows.Forms.Label();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnVer = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnGenerar = new System.Windows.Forms.Button();
			this.btnSeleccionar = new System.Windows.Forms.Button();
			this.btnActualizar = new System.Windows.Forms.Button();
			this.lblContador = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.uGridOperaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridOperaciones
			// 
			this.uGridOperaciones.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridOperaciones.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridOperaciones.DisplayLayout.Appearance = appearance1;
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridOperaciones.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridOperaciones.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridOperaciones.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.uGridOperaciones.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridOperaciones.DisplayLayout.Override.CellAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 8F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridOperaciones.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.uGridOperaciones.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridOperaciones.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridOperaciones.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.uGridOperaciones.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridOperaciones.Location = new System.Drawing.Point(8, 56);
			this.uGridOperaciones.Name = "uGridOperaciones";
			this.uGridOperaciones.Size = new System.Drawing.Size(976, 304);
			this.uGridOperaciones.TabIndex = 737;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.BackColor = System.Drawing.Color.Transparent;
			this.lblBodega.Location = new System.Drawing.Point(8, 10);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(31, 17);
			this.lblBodega.TabIndex = 759;
			this.lblBodega.Text = "Corte";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtDesde
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance7;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton1);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(56, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 758;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(184, 6);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 760;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
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
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 40);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(992, 8);
			this.groupBox1.TabIndex = 761;
			this.groupBox1.TabStop = false;
			// 
			// btnGenerar
			// 
			this.btnGenerar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGenerar.CausesValidation = false;
			this.btnGenerar.Enabled = false;
			this.btnGenerar.Image = ((System.Drawing.Image)(resources.GetObject("btnGenerar.Image")));
			this.btnGenerar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGenerar.Location = new System.Drawing.Point(896, 8);
			this.btnGenerar.Name = "btnGenerar";
			this.btnGenerar.Size = new System.Drawing.Size(88, 23);
			this.btnGenerar.TabIndex = 762;
			this.btnGenerar.Text = "Generar TXT";
			this.btnGenerar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGenerar.Click += new System.EventHandler(this.btnGenerar_Click);
			// 
			// btnSeleccionar
			// 
			this.btnSeleccionar.CausesValidation = false;
			this.btnSeleccionar.Image = ((System.Drawing.Image)(resources.GetObject("btnSeleccionar.Image")));
			this.btnSeleccionar.Location = new System.Drawing.Point(8, 56);
			this.btnSeleccionar.Name = "btnSeleccionar";
			this.btnSeleccionar.Size = new System.Drawing.Size(25, 25);
			this.btnSeleccionar.TabIndex = 763;
			this.btnSeleccionar.Click += new System.EventHandler(this.btnSeleccionar_Click);
			// 
			// btnActualizar
			// 
			this.btnActualizar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnActualizar.CausesValidation = false;
			this.btnActualizar.Enabled = false;
			this.btnActualizar.Image = ((System.Drawing.Image)(resources.GetObject("btnActualizar.Image")));
			this.btnActualizar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnActualizar.Location = new System.Drawing.Point(784, 8);
			this.btnActualizar.Name = "btnActualizar";
			this.btnActualizar.Size = new System.Drawing.Size(104, 23);
			this.btnActualizar.TabIndex = 764;
			this.btnActualizar.Text = "Actualizar T A";
			this.btnActualizar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnActualizar.Click += new System.EventHandler(this.btnActualizar_Click);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.BackColor = System.Drawing.Color.Transparent;
			this.lblContador.Location = new System.Drawing.Point(368, 8);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 17);
			this.lblContador.TabIndex = 765;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// frmCre_CarteraRefinanciadaPichincha
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(992, 366);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.btnActualizar);
			this.Controls.Add(this.btnSeleccionar);
			this.Controls.Add(this.btnGenerar);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.uGridOperaciones);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.Name = "frmCre_CarteraRefinanciadaPichincha";
			this.Text = "Cartera Refinanciada Pichincha";
			this.Load += new System.EventHandler(this.frmCre_CarteraRefinanciadaPichincha_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridOperaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void UnloadMe()
		{
			this.Close();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		CultureInfo us = new CultureInfo("en-US");
		
		private void frmCre_CarteraRefinanciadaPichincha_Load(object sender, System.EventArgs e)
		{		
			miAcceso = new Acceso(cdsSeteoF, "08320205");
			if (!Funcion.Permiso("1515", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Seguro DFI", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			if (miAcceso.BActualizar)this.btnActualizar.Enabled = true;
			if (miAcceso.BGenerar)this.btnGenerar.Enabled = true;

			this.dtDesde.Value = DateTime.Today;
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.dsGeneral("Exec Cre_ListaClientesRefinanciamiento", this.uGridOperaciones);
			this.lblContador.Text = this.uGridOperaciones.Rows.Count + " Registros Encontrados";
		}

		private void btnGenerar_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.OK == MessageBox.Show("Esta seguro de GENERAR EL ARCHIVO", "Point Technology", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				try
				{					
					cdsSeteoF.Schema.Connections[0].Open();
					C1.Data.SchemaObjects.Connection miConeccion = cdsSeteoF.Schema.Connections[0];
					miConeccion = cdsSeteoF.Schema.Connections[0];

					SaveFileDialog sArchivo = new SaveFileDialog();

					sArchivo.Filter = "Documentos de texto (*.txt)|*.txt";
					sArchivo.FileName = "Ejemplo.txt";
					sArchivo.OverwritePrompt = true;
					sArchivo.Title = "Genera Archivo Diferimiento Banco Pichincha";

					if (sArchivo.ShowDialog() == DialogResult.OK)
					{
						using (StreamWriter file = new StreamWriter(sArchivo.FileName, true))
						{											
							int i = 0;
							int iOperaciones = 0;

							foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridOperaciones.Rows)
							{
								#region Variables
								int iCuotaInicial = 0;
								int iCuotasDiferimiento = 0;
								int iDiasCalculo = 0;
								decimal dInteresCalculado = 0.00m;
								decimal dValorCuotaDiferida = 0.00m;
								decimal dCapitalCuotaDiferida = 0.00m;
								decimal dValorCuotaNuevo = 0.00m;
								decimal dValorRefinanciamiento = 0.00m;
								decimal dValorMensualRefinanciamiento = 0.00m;
								decimal dCapitalVendido = 0.00m;
								decimal dValorDifRef = 0.00m;
								#endregion Variables

								if (Convert.ToBoolean(dr.Cells["Sel"].Value))
								{
									#region Calculo Valores Por Diferimiento
									string sSQL = string.Format("Exec Cre_CalculaDiferimientoCuotasPichincha {0}", (int)dr.Cells["idCompra"].Value);
									SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
									drC.Read();
									if (drC.HasRows)
									{	
										iCuotaInicial = drC.GetInt32(0);
										iCuotasDiferimiento = drC.GetInt32(1);
										iDiasCalculo = drC.GetInt32(2);
										dInteresCalculado = drC.GetDecimal(3);
										dValorCuotaDiferida = drC.GetDecimal(4);
										dCapitalCuotaDiferida = drC.GetDecimal(5);
										dValorCuotaNuevo = drC.GetDecimal(6);
										dValorRefinanciamiento = drC.GetDecimal(7);
										dValorMensualRefinanciamiento = drC.GetDecimal(8);
										dCapitalVendido = drC.GetDecimal(9);
										dValorDifRef = drC.GetDecimal(10);
									}
									drC.Close();
									#endregion Calculo Valores Por Diferimiento

									#region Genera Archivo
									DataSet oDSCompras = new DataSet();
									string sSQLCompras = string.Format("Exec Cre_TablaAmortizacionDifPichincha {0}", (int)dr.Cells["idCompra"].Value);
									SqlDataAdapter oDACompras = new SqlDataAdapter(sSQLCompras, (SqlConnection) miConeccion.DbConnection);
									oDACompras.Fill(oDSCompras);

									if (oDSCompras.Tables[0].Rows.Count > 0)
									{
										#region Enabezados
										if (i == 0)
										{
											string iLinea = "TACDIF" + "\t" +	"TAFECO" + "\t" +	"TAFESA" + "\t" +	"TANUOP" + "\t" +	"TAOPAL" + "\t" +	"TAFEVC" + "\t" +	"TANUDV" + "\t" +	"TATPRB" + "\t" +	"TACDRB" + "\t" +	"TATASA" + "\t" +	"TAVALO" + "\t" +	"TASALD" + "\t" +	"TAVARE" + "\t" +	"TAPAGO";
											file.WriteLine(iLinea);
										}
										#endregion Enabezados

										foreach(DataRow oDRCompras in oDSCompras.Tables[0].Rows)
										{
											#region Lectura Datos Tabla De Amortizacion
											string sTACDIF = oDRCompras["TACDIF"].ToString();
											string sTAFECO = oDRCompras["TAFECO"].ToString();
											string sTAFESA = oDRCompras["TAFESA"].ToString();
											string sTANUOP = oDRCompras["TANUOP"].ToString();
											string sTAOPAL = oDRCompras["TAOPAL"].ToString();
											DateTime sTAFEVC = Convert.ToDateTime(oDRCompras["TAFEVC"]);//.ToString();
											int iNumCuotas = (int)oDRCompras["NumCuotas"];
											int iCuotas_Vendidas = (int)oDRCompras["Cuotas_Vendidas"];
											int iNumeroCuota = (int)oDRCompras["NumeroCuota"];
											DateTime dtVence = (DateTime)oDRCompras["Vence"];
											decimal dCapital = (decimal)oDRCompras["Capital"];
											decimal dABCapital = (decimal)oDRCompras["ABCapital"];
											decimal dInteres = (decimal)oDRCompras["Interes"];
											decimal dABInteres = (decimal)oDRCompras["ABInteres"];
											decimal dABMora = (decimal)oDRCompras["ABMora"];
											decimal dSaldoCapital = (decimal)oDRCompras["SaldoCapital"];
											decimal dTasaInteres = (decimal)oDRCompras["TasaInteres"];
											int iEstado = (int)oDRCompras["Estado"];
											bool bUltima = (bool)oDRCompras["Ultima"];
											
											decimal dSaldoCapitalCuota = dCapital - dABCapital;
											decimal dSaldoInteresCuota = dInteres - dABInteres;
											decimal dValorCuota = dCapital + dInteres;
											decimal dValorPagadoCuota = dABCapital + dABInteres;
											decimal dSaldoCuota = dValorCuota - dValorPagadoCuota;
											string sEstado = "CAN";
											#endregion Lectura Datos Tabla De Amortizacion
											
											if (iNumeroCuota < iCuotaInicial)
											{
												#region Cuotas Pagadas
												int iDifCuotas = iNumCuotas - iCuotas_Vendidas;
												iNumeroCuota = iNumeroCuota - iDifCuotas;

												/* CAPITAL */
												string iLinea = sTACDIF + "\t" + sTAFECO + "\t" + sTAFESA + "\t" + sTANUOP + "\t" + sTAOPAL + "\t" + sTAFEVC.ToString("yyyyMMdd") + "\t" 
													+ iNumeroCuota.ToString() + "\t" + "CAP" + "\t" + "PRI" + "\t" 
													+ "0.00" + "\t" + dCapital.ToString() + "\t" + dSaldoCapitalCuota.ToString("0.00") + "\t" + dSaldoCapital.ToString("0.00") + "\t" + dABCapital.ToString();

												file.WriteLine(iLinea);

												/* INTERES */
												iLinea = sTACDIF + "\t" + sTAFECO + "\t" + sTAFESA + "\t" + sTANUOP + "\t" + sTAOPAL + "\t" + sTAFEVC.ToString("yyyyMMdd") + "\t" 
													+ iNumeroCuota.ToString() + "\t" + "INT" + "\t" + "VCT" + "\t" 
													+ dTasaInteres.ToString() + "\t" + dInteres.ToString() + "\t" + dSaldoInteresCuota.ToString("0.00") + "\t" + "0.00" + "\t" + dABInteres.ToString();

												file.WriteLine(iLinea);

												/* MORA */
												if (dABMora > 0.00m)
												{
													iLinea = sTACDIF + "\t" + sTAFECO + "\t" + sTAFESA + "\t" + sTANUOP + "\t" + sTAOPAL + "\t" + sTAFEVC.ToString("yyyyMMdd") + "\t" 
														+ iNumeroCuota.ToString() + "\t" + "MOR" + "\t" + "NOR" + "\t" 
														+ "0.00" + "\t" + dABMora.ToString() + "\t" + "0.00" + "\t" + "0.00" + "\t" + dABMora.ToString();

													file.WriteLine(iLinea);
												}
												
												/* ESTADO */
												iLinea = sTACDIF + "\t" + sTAFECO + "\t" + sTAFESA + "\t" + sTANUOP + "\t" + sTAOPAL + "\t" + sTAFEVC.ToString("yyyyMMdd") + "\t" 
													+ iNumeroCuota.ToString() + "\t" + "CTI" + "\t" + sEstado + "\t" 
													+ "0.00" + "\t" + dValorCuota.ToString() + "\t" + dSaldoCuota.ToString("0.00") + "\t" + "0.00" + "\t" + dValorPagadoCuota.ToString();

												file.WriteLine(iLinea);

												dCapitalVendido = dCapitalVendido - dCapital;
											
												//continue;
												#endregion Cuotas Pagadas
											}
											else if (iNumeroCuota == iCuotaInicial)
											{
												#region Cuota Diferida
												int iDifCuotas = iNumCuotas - iCuotas_Vendidas;
												iNumeroCuota = iNumeroCuota - iDifCuotas + iCuotasDiferimiento;
												dCapitalVendido = dCapitalVendido - dCapital;//CuotaDiferida;
												sTAFEVC = sTAFEVC.AddMonths(4);

												/* CAPITAL dCapitalCuotaDiferida ES REEMPLAZADA POR DCAPITAL*/
												string iLinea = sTACDIF + "\t" + sTAFECO + "\t" + sTAFESA + "\t" + sTANUOP + "\t" + sTAOPAL + "\t" + sTAFEVC.ToString("yyyyMMdd") + "\t" 
													+ iNumeroCuota.ToString() + "\t" + "CAP" + "\t" + "PRI" + "\t" 
													+ "0.00" + "\t" + dCapital.ToString() + "\t" + dCapital.ToString() + "\t" + dCapitalVendido.ToString() + "\t" + "0.00";

												file.WriteLine(iLinea);

												/* INTERES */
												iLinea = sTACDIF + "\t" + sTAFECO + "\t" + sTAFESA + "\t" + sTANUOP + "\t" + sTAOPAL + "\t" + sTAFEVC.ToString("yyyyMMdd") + "\t" 
													+ iNumeroCuota.ToString() + "\t" + "INT" + "\t" + "VCT" + "\t" 
													+ dTasaInteres.ToString() + "\t" + dInteresCalculado.ToString() + "\t" + dInteresCalculado.ToString() + "\t" + "0.00" + "\t" + "0.00";

												file.WriteLine(iLinea);

												/* VALOR REFINANCIAMIENTO */
												iLinea = sTACDIF + "\t" + sTAFECO + "\t" + sTAFESA + "\t" + sTANUOP + "\t" + sTAOPAL + "\t" + sTAFEVC.ToString("yyyyMMdd") + "\t" 
													+ iNumeroCuota.ToString() + "\t" + "GAS" + "\t" + "QKP" + "\t" 
													+ "0.00" + "\t" + dValorRefinanciamiento.ToString() + "\t" + dValorRefinanciamiento.ToString() + "\t" + "0.00" + "\t" + "0.00";

												file.WriteLine(iLinea);
																								
												/* ESTADO */
												iLinea = sTACDIF + "\t" + sTAFECO + "\t" + sTAFESA + "\t" + sTANUOP + "\t" + sTAOPAL + "\t" + sTAFEVC.ToString("yyyyMMdd") + "\t" 
													+ iNumeroCuota.ToString() + "\t" + "CTI" + "\t" + "PND" + "\t" 
													+ "0.00" + "\t" + dValorCuotaDiferida.ToString() + "\t" + dValorCuotaDiferida.ToString() + "\t" + "0.00" + "\t" + "0.00";

												file.WriteLine(iLinea);
											
												//continue;
												#endregion Cuota Diferida
											}
											else if (iNumeroCuota > iCuotaInicial)
											{
												#region Cuota No Devenga
												int iDifCuotas = iNumCuotas - iCuotas_Vendidas;
												iNumeroCuota = iNumeroCuota - iDifCuotas + iCuotasDiferimiento;
												//int iCuotasPend = iCuotas_Vendidas + iCuotasDiferimiento - iNumeroCuota;
												sTAFEVC = sTAFEVC.AddMonths(4);
												
												//												decimal dInteresCuotaNueva = (dCapitalVendido * 16.06m/100.00m)/360 * 30;												
												//												decimal dCapitalCuotaNueva = (dValorCuotaNuevo - dValorMensualRefinanciamiento) - decimal.Round(dInteresCuotaNueva, 2);
												dCapitalVendido = dCapitalVendido - dCapital;//CuotaNueva;

												if (bUltima) dValorMensualRefinanciamiento = dValorDifRef;

												/* CAPITAL  dCapitalCuotaNueva ES REEMPLAZADA POR DCAPITAL */
												string iLinea = sTACDIF + "\t" + sTAFECO + "\t" + sTAFESA + "\t" + sTANUOP + "\t" + sTAOPAL + "\t" + sTAFEVC.ToString("yyyyMMdd") + "\t" 
													+ iNumeroCuota.ToString() + "\t" + "CAP" + "\t" + "PRI" + "\t" 
													+ "0.00" + "\t" + dCapital.ToString() + "\t" + dCapital.ToString() + "\t" + dCapitalVendido.ToString("0.00") + "\t" + "0.00";

												file.WriteLine(iLinea);

												/* INTERES dInteresCuotaNueva ES REEMPLAZADA POR DINTERES */
												iLinea = sTACDIF + "\t" + sTAFECO + "\t" + sTAFESA + "\t" + sTANUOP + "\t" + sTAOPAL + "\t" + sTAFEVC.ToString("yyyyMMdd") + "\t" 
													+ iNumeroCuota.ToString() + "\t" + "INT" + "\t" + "VCT" + "\t" 
													+ dTasaInteres.ToString() + "\t" + decimal.Round(dInteres, 2).ToString() + "\t" + decimal.Round(dInteres, 2).ToString()+ "\t" + "0.00" + "\t" + "0.00";

												file.WriteLine(iLinea);

												/* VALOR MENSUAL REFINANCIAMIENTO */
												iLinea = sTACDIF + "\t" + sTAFECO + "\t" + sTAFESA + "\t" + sTANUOP + "\t" + sTAOPAL + "\t" + sTAFEVC.ToString("yyyyMMdd") + "\t" 
													+ iNumeroCuota.ToString() + "\t" + "GAS" + "\t" + "QKP" + "\t" 
													+ "0.00" + "\t" + dValorMensualRefinanciamiento.ToString() + "\t" + dValorMensualRefinanciamiento.ToString() + "\t" + "0.00" + "\t" + "0.00";

												file.WriteLine(iLinea);
		
												dValorCuota = dValorCuota + dValorMensualRefinanciamiento;

												/* ESTADO dValorCuotaNuevo ES REEMPLAZDA POR dValorCuota */
												iLinea = sTACDIF + "\t" + sTAFECO + "\t" + sTAFESA + "\t" + sTANUOP + "\t" + sTAOPAL + "\t" + sTAFEVC.ToString("yyyyMMdd") + "\t" 
													+ iNumeroCuota.ToString() + "\t" + "CTI" + "\t" + "AFE" + "\t" 
													+ "0.00" + "\t" + dValorCuota.ToString() + "\t" + dValorCuota.ToString() + "\t" + "0.00" + "\t" + "0.00";

												file.WriteLine(iLinea);
												#endregion Cuota No Devenga
											}
											
											i++;
										}	
	
										iOperaciones++;
									}
									#endregion Genera Archivo
								}														
							}

							MessageBox.Show(string.Format("Archivo Generado Correctamente, {0} Operaciones, {1} Cuotas", iOperaciones, i), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						}
					}
				}
				catch (Exception exc)
				{
					MessageBox.Show(exc.Message);
				}
			}
		}

		private void btnSeleccionar_Click(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridOperaciones.Rows.All)			
				dr.Cells["Sel"].Value = !(bool)dr.Cells["Sel"].Value;
		}

		private void btnActualizar_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.OK == MessageBox.Show("Esta seguro de GENERAR EL ARCHIVO", "Point Technology", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				try
				{					
					cdsSeteoF.Schema.Connections[0].Open();
					C1.Data.SchemaObjects.Connection miConeccion = cdsSeteoF.Schema.Connections[0];
					miConeccion = cdsSeteoF.Schema.Connections[0];
			
					int i = 0;
					int iOperaciones = 0;
			
					foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridOperaciones.Rows)
					{
						if (Convert.ToBoolean(dr.Cells["Sel"].Value))
						{
							#region Variables
							int iCuotaInicial = 0;
							int iCuotasDiferimiento = 0;
							int iDiasCalculo = 0;
							decimal dInteresCalculado = 0.00m;
							decimal dValorCuotaDiferida = 0.00m;
							decimal dCapitalCuotaDiferida = 0.00m;
							decimal dValorCuotaNuevo = 0.00m;
							decimal dValorRefinanciamiento = 0.00m;
							decimal dValorMensualRefinanciamiento = 0.00m;
							decimal dCapitalVendido = 0.00m;
							decimal dValorDifRef = 0.00m;							
							#endregion Variables
			
							#region Calculo Valores Por Diferimiento
							string sSQL = string.Format("Exec Cre_CalculaDiferimientoCuotasPichincha {0}", (int)dr.Cells["idCompra"].Value);
							SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
							drC.Read();
							if (drC.HasRows)
							{	
								iCuotaInicial = drC.GetInt32(0);
								iCuotasDiferimiento = drC.GetInt32(1);
								iDiasCalculo = drC.GetInt32(2);
								dInteresCalculado = drC.GetDecimal(3);
								dValorCuotaDiferida = drC.GetDecimal(4);
								dCapitalCuotaDiferida = drC.GetDecimal(5);
								dValorCuotaNuevo = drC.GetDecimal(6);
								dValorRefinanciamiento = drC.GetDecimal(7);
								dValorMensualRefinanciamiento = drC.GetDecimal(8);
								dCapitalVendido = drC.GetDecimal(9);
								dValorDifRef = drC.GetDecimal(10);
							}
							drC.Close();
							#endregion Calculo Valores Por Diferimiento
			
							#region Genera Archivo
							DataSet oDSCompras = new DataSet();
							string sSQLCompras = string.Format("Exec Cre_TablaAmortizacionDifPichincha {0}", (int)dr.Cells["idCompra"].Value);
							SqlDataAdapter oDACompras = new SqlDataAdapter(sSQLCompras, (SqlConnection) miConeccion.DbConnection);
							oDACompras.Fill(oDSCompras);
			
							if (oDSCompras.Tables[0].Rows.Count > 0)
							{													
								foreach(DataRow oDRCompras in oDSCompras.Tables[0].Rows)
								{
									#region Lectura Datos Tabla De Amortizacion
									int idCre_TablaDeAmortizacion = (int)oDRCompras["idCre_TablaDeAmortizacion"];
									string sTACDIF = oDRCompras["TACDIF"].ToString();
									string sTAFECO = oDRCompras["TAFECO"].ToString();
									string sTAFESA = oDRCompras["TAFESA"].ToString();
									string sTANUOP = oDRCompras["TANUOP"].ToString();
									string sTAOPAL = oDRCompras["TAOPAL"].ToString();
									DateTime sTAFEVC = Convert.ToDateTime(oDRCompras["TAFEVC"]);//string sTAFEVC = oDRCompras["TAFEVC"].ToString();
									int iNumCuotas = (int)oDRCompras["NumCuotas"];											
									int iCuotas_Vendidas = (int)oDRCompras["Cuotas_Vendidas"];
									int iNumeroCuota = (int)oDRCompras["NumeroCuota"];
									DateTime dtInicia = (DateTime)oDRCompras["Inicia"];
									DateTime dtVence = (DateTime)oDRCompras["Vence"];
									decimal dCapital = (decimal)oDRCompras["Capital"];
									decimal dABCapital = (decimal)oDRCompras["ABCapital"];
									decimal dInteres = (decimal)oDRCompras["Interes"];
									decimal dABInteres = (decimal)oDRCompras["ABInteres"];
									decimal dABMora = (decimal)oDRCompras["ABMora"];
									decimal dSaldoCapital = (decimal)oDRCompras["SaldoCapital"];
									decimal dTasaInteres = (decimal)oDRCompras["TasaInteres"];
									int iEstado = (int)oDRCompras["Estado"];
									bool bUltima = (bool)oDRCompras["Ultima"];
														
									decimal dSaldoCapitalCuota = dCapital - dABCapital;
									decimal dSaldoInteresCuota = dInteres - dABInteres;
									decimal dValorCuota = dCapital + dInteres;
									decimal dValorPagadoCuota = dABCapital + dABInteres;
									decimal dSaldoCuota = dValorCuota - dValorPagadoCuota;
									string sEstado = "CAN";											
									#endregion Lectura Datos Tabla De Amortizacion
												
									if (iNumeroCuota == iCuotaInicial)
									{
										#region Cuota Diferida
//										int iDifCuotas = iNumCuotas - iCuotas_Vendidas;
//										iNumeroCuota = iNumeroCuota - iDifCuotas;// + iCuotasDiferimiento;
			
										sSQL = string.Format("Insert Into Cre_TablaDeAmortizacion (idCompra, NumeroCuota, Inicia, Vence, FechaCancelacion, ValorCuota, Interes, Capital, ABCapital, Abono, Saldo, Estado, NumeroLote) Values ({0}, {1}, '{2}', '{3}', '20200530', 0, 0, 0, 0, 0, 0, 4, {4})",
											(int)dr.Cells["idCompra"].Value, iNumeroCuota, dtInicia.ToString("yyyyMMdd"), dtVence.ToString("yyyyMMdd"), (int)dr.Cells["NumeroLote"].Value);
										Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);									
																
										iNumeroCuota++;
			
										//dtInicia = dtInicia.AddMonths(1);
										dtVence = dtVence.AddMonths(1);
			
										sSQL = string.Format("Insert Into Cre_TablaDeAmortizacion (idCompra, NumeroCuota, Inicia, Vence, FechaCancelacion, ValorCuota, Interes, Capital, ABCapital, Abono, Saldo, Estado, NumeroLote) Values ({0}, {1}, '{2}', '{3}', '20200530', 0, 0, 0, 0, 0, 0, 4, {4})",
											(int)dr.Cells["idCompra"].Value, iNumeroCuota, dtInicia.ToString("yyyyMMdd"), dtVence.ToString("yyyyMMdd"), (int)dr.Cells["NumeroLote"].Value);
										Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);									
																
										iNumeroCuota++;
			
										//dtInicia = dtInicia.AddMonths(1);
										dtVence = dtVence.AddMonths(1);
			
										sSQL = string.Format("Insert Into Cre_TablaDeAmortizacion (idCompra, NumeroCuota, Inicia, Vence, FechaCancelacion, ValorCuota, Interes, Capital, ABCapital, Abono, Saldo, Estado, Numerolote) Values ({0}, {1}, '{2}', '{3}', '20200530', 0, 0, 0, 0, 0, 0, 4, {4})",
											(int)dr.Cells["idCompra"].Value, iNumeroCuota, dtInicia.ToString("yyyyMMdd"), dtVence.ToString("yyyyMMdd"), (int)dr.Cells["NumeroLote"].Value);
										Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
																
										iNumeroCuota++;
			
										//dtInicia = dtInicia.AddMonths(1);
										dtVence = dtVence.AddMonths(1);
			
										sSQL = string.Format("Insert Into Cre_TablaDeAmortizacion (idCompra, NumeroCuota, Inicia, Vence, FechaCancelacion, ValorCuota, Interes, Capital, ABCapital, Abono, Saldo, Estado, NumeroLote) Values ({0}, {1}, '{2}', '{3}', '20200530', 0, 0, 0, 0, 0, 0, 4, {4})",
											(int)dr.Cells["idCompra"].Value, iNumeroCuota, dtInicia.ToString("yyyyMMdd"), dtVence.ToString("yyyyMMdd"), (int)dr.Cells["NumeroLote"].Value);
										Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
										#endregion Cuotas Diferidas
				
										iNumeroCuota++;
										//dtInicia = dtInicia.AddMonths(1);
										dtVence = dtVence.AddMonths(1);
			
										/* CAPITAL */

										sSQL = string.Format("Update Cre_TablaDeAmortizacion Set NumeroCuota = {0}, Inicia = '{1}', Vence = '{2}', Interes = {3}, Diferimiento = {4}, ValorCuota = Capital + {3} Where idCre_TablaDeAmortizacion = {5}",
											iNumeroCuota, dtInicia.ToString("yyyyMMdd"), dtVence.ToString("yyyyMMdd"), dInteresCalculado, dValorRefinanciamiento, idCre_TablaDeAmortizacion);
										Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
									}
									else if (iNumeroCuota > iCuotaInicial)
									{
										#region Cuota No Devenga
//										int iDifCuotas = iNumCuotas - iCuotas_Vendidas;
//										iNumeroCuota = iNumeroCuota - iDifCuotas + 4;
										iNumeroCuota = iNumeroCuota + 4;
														
										dtInicia = dtInicia.AddMonths(4);
										dtVence = dtVence.AddMonths(4);
			
										if (bUltima) dValorMensualRefinanciamiento = dValorDifRef;
			
										sSQL = string.Format("Update Cre_TablaDeAmortizacion Set NumeroCuota = {0}, Inicia = '{1}', Vence = '{2}', Diferimiento = {3} Where idCre_TablaDeAmortizacion = {4}",
											iNumeroCuota, dtInicia.ToString("yyyyMMdd"), dtVence.ToString("yyyyMMdd"), dValorMensualRefinanciamiento, idCre_TablaDeAmortizacion);
										Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
										#endregion Cuota No Devenga
									}
														
									i++;											
								}									
													
								iOperaciones++;
							}
							#endregion Genera Archivo
						}													
					}
			
					MessageBox.Show(string.Format("Archivo Generado Correctamente, {0} Operaciones, {1} Cuotas", iOperaciones, i), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);						
				}
				catch (Exception exc)
				{
					MessageBox.Show(exc.Message);
				}
			}
		}
	}
}