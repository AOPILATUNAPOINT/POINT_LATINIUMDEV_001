using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.IO;
using System.Text;
using System.Threading;
using System.Data;
using System.Data.SqlClient;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;
using C1.Data;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteCashManagement.
	/// </summary>
	public class frmReporteCashManagement : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Button btnGenerar;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnImportar;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReporteCashManagement()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Contrapartida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Moneda");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaDeCobro");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Referencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoId");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroId");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NombreBeneficiario");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 3, true);
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Contrapartida");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Moneda");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FormaDeCobro");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Referencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoId");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroId");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NombreBeneficiario");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteCashManagement));
			this.uGridInformacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.btnVer = new System.Windows.Forms.Button();
			this.btnExcel = new System.Windows.Forms.Button();
			this.btnGenerar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnImportar = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridInformacion
			// 
			this.uGridInformacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridInformacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridInformacion.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridInformacion.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 70;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn2.Header.Caption = "Contra Partida";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 120;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn3.CellAppearance = appearance2;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 70;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance3;
			ultraGridColumn4.Format = "#,##0.00";
			ultraGridColumn4.Header.VisiblePosition = 8;
			ultraGridColumn4.Width = 100;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn5.CellAppearance = appearance4;
			ultraGridColumn5.Header.Caption = "Forma De Cobro";
			ultraGridColumn5.Header.VisiblePosition = 3;
			ultraGridColumn5.Width = 100;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn6.Header.VisiblePosition = 4;
			ultraGridColumn6.Width = 150;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn7.CellAppearance = appearance5;
			ultraGridColumn7.Header.Caption = "Tipo Id";
			ultraGridColumn7.Header.VisiblePosition = 5;
			ultraGridColumn7.Width = 60;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Left;
			ultraGridColumn8.CellAppearance = appearance6;
			ultraGridColumn8.Header.Caption = "Cedula";
			ultraGridColumn8.Header.VisiblePosition = 6;
			ultraGridColumn8.Width = 100;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn9.Header.Caption = "Cliente";
			ultraGridColumn9.Header.VisiblePosition = 7;
			ultraGridColumn9.Width = 300;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			ultraGridBand1.Header.Caption = "Periodo A";
			ultraGridBand1.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			ultraGridBand1.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance7;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridInformacion.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridInformacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridInformacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridInformacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance8.BackColor = System.Drawing.Color.Transparent;
			this.uGridInformacion.DisplayLayout.Override.CardAreaAppearance = appearance8;
			appearance9.ForeColor = System.Drawing.Color.Black;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInformacion.DisplayLayout.Override.CellAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance10.FontData.BoldAsString = "True";
			appearance10.FontData.Name = "Arial";
			appearance10.FontData.SizeInPoints = 8F;
			appearance10.ForeColor = System.Drawing.Color.White;
			appearance10.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInformacion.DisplayLayout.Override.HeaderAppearance = appearance10;
			this.uGridInformacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowSelectorAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.SelectedRowAppearance = appearance12;
			this.uGridInformacion.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInformacion.Location = new System.Drawing.Point(8, 56);
			this.uGridInformacion.Name = "uGridInformacion";
			this.uGridInformacion.Size = new System.Drawing.Size(1091, 304);
			this.uGridInformacion.TabIndex = 767;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn4.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9});
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Location = new System.Drawing.Point(0, 40);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1301, 8);
			this.groupBox2.TabIndex = 770;
			this.groupBox2.TabStop = false;
			// 
			// btnVer
			// 
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(8, 8);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(128, 24);
			this.btnVer.TabIndex = 771;
			this.btnVer.Text = "Generar Información";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Enabled = false;
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(1008, 8);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(96, 24);
			this.btnExcel.TabIndex = 772;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnGenerar
			// 
			this.btnGenerar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGenerar.Enabled = false;
			this.btnGenerar.Image = ((System.Drawing.Image)(resources.GetObject("btnGenerar.Image")));
			this.btnGenerar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGenerar.Location = new System.Drawing.Point(808, 8);
			this.btnGenerar.Name = "btnGenerar";
			this.btnGenerar.Size = new System.Drawing.Size(96, 24);
			this.btnGenerar.TabIndex = 773;
			this.btnGenerar.Text = "Generar TXT";
			this.btnGenerar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGenerar.Click += new System.EventHandler(this.btnGenerar_Click);
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
			// btnImportar
			// 
			this.btnImportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnImportar.CausesValidation = false;
			this.btnImportar.Enabled = false;
			this.btnImportar.Image = ((System.Drawing.Image)(resources.GetObject("btnImportar.Image")));
			this.btnImportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImportar.Location = new System.Drawing.Point(912, 8);
			this.btnImportar.Name = "btnImportar";
			this.btnImportar.Size = new System.Drawing.Size(88, 24);
			this.btnImportar.TabIndex = 774;
			this.btnImportar.Text = "Importar";
			this.btnImportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImportar.Click += new System.EventHandler(this.btnImportar_Click);
			// 
			// frmReporteCashManagement
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1107, 366);
			this.ControlBox = false;
			this.Controls.Add(this.btnImportar);
			this.Controls.Add(this.btnGenerar);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.uGridInformacion);
			this.KeyPreview = true;
			this.Name = "frmReporteCashManagement";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Cash Management";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReporteCashManagement_KeyDown);
			this.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.frmReporteCashManagement_KeyPress);
			this.Load += new System.EventHandler(this.frmReporteCashManagement_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion


		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}

		private void frmReporteCashManagement_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			
		}

		private void btnGenerar_Click(object sender, System.EventArgs e)
		{      
			SaveFileDialog sArchivo = new SaveFileDialog();

			sArchivo.Filter = "Documentos de texto (*.txt)|*.txt";
			sArchivo.FileName = "CashManagemnet.txt";
			sArchivo.OverwritePrompt = true;
			sArchivo.Title = "Genera Archivo de Pagos Banco Pichincha";
				                                
			if (sArchivo.ShowDialog() == DialogResult.OK)
			{					
				#region Pagos
				using (StreamWriter file = new StreamWriter(sArchivo.FileName, true))
				{
					#region Normales
					int i = 0;
					foreach (DataRow dr in FuncionesProcedimientos.dtGeneral(string.Format("Exec Cob_ListaCashManagement")).Rows)
					{
						try
						{
							int iValor = int.Parse(dr["Valor"].ToString().Replace(".", ""));

							string iLinea = dr["Tipo"].ToString() + "\t" + dr["ContraPartida"].ToString() + "\t" + dr["Moneda"].ToString() + "\t" + iValor.ToString() + "\t"
								+ dr["FormaDeCobro"].ToString() + "\t" + " " + "\t" + " "+ "\t"+ dr["Referencia"].ToString() + "\t" + dr["TipoId"].ToString() + "\t"
								+ dr["NumeroId"].ToString() + "\t" + dr["NombreBeneficiario"].ToString() ;
							
							file.WriteLine(iLinea);

							i++;
						}						
						catch (Exception exc)
						{
							MessageBox.Show(exc.Message + " Linea : " + i.ToString());
							continue;
						}
					}

					if (i > 0)
						MessageBox.Show(string.Format("Archivo generado correctamente con {0} lineas", i), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					#endregion Normales
        }
				#endregion Pagos
			}
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Exec Cob_ListaCashManagement");
			this.uGridInformacion.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);
		}

		private void frmReporteCashManagement_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}
		
		private void frmReporteCashManagement_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "082915");
			if (!Funcion.Permiso("1176", cdsSeteoF))
			{				
				MessageBox.Show("No tiene Acceso a CASH MANAGEMENT", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			if (miAcceso.BExportar) this.btnExcel.Enabled =  true;
			if (miAcceso.BGenerar) this.btnGenerar.Enabled =  true;
			if (miAcceso.BImportar) this.btnImportar.Enabled = true;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnImportar_Click(object sender, System.EventArgs e)
		{			
			#region Variables
			int iCont = 0;
			int IdHistorial = 0;
			string sCedula = "";
			string sNombre = "";
			decimal dValor = 0.00m;
			string sFecha= "";
			string sReferencia = "";						
			string sTipo = "";
			int idCompra = 0;
			int idAnticipo = 0;
			#endregion Variables

			#region Sube archivo
			try
			{
				OpenFileDialog sArchivo = new OpenFileDialog();

				sArchivo.Filter = "Microsoft Excel (*.xls)|*.xls";

				if (sArchivo.ShowDialog() == System.Windows.Forms.DialogResult.OK && sArchivo.FileName.Length > 0)
				{
					this.Cursor = Cursors.WaitCursor;
										
					foreach (DataRow row in FuncionesProcedimientos.ImportaDeExcel(sArchivo.FileName, "SELECT * FROM [PAGOS$]", "PAGOS").Tables[0].Rows)
					{						
						if (row["Cedula"] != System.DBNull.Value) sCedula = row["Cedula"].ToString();
						if (row["Nombre"] != System.DBNull.Value) sNombre = row["Nombre"].ToString();
						if (row["Fecha"] != System.DBNull.Value) sFecha = row["Fecha"].ToString();;
						if (row["Referencia"] != System.DBNull.Value) sReferencia = row["Referencia"].ToString();
						//if (row["Tipo pago"] != System.DBNull.Value) sTipoPago = row["Tipo pago"].ToString();
						if (row["Valor"] != System.DBNull.Value) dValor = decimal.Parse(row["Valor"].ToString());
						if (row["Medio"] != System.DBNull.Value) sTipo = row["Medio"].ToString();						

						string sSQLAct = string.Format("Exec GrabaHistorialCashManagment {0}, '{1}', '{2}', {3}, '{4}', '{5}','{6}',{7} ", 
							IdHistorial, sCedula, sNombre, dValor, Convert.ToDateTime(sFecha).ToString("yyyyMMdd"), sReferencia,  sTipo, 0);						
						Funcion.EjecutaSQL(cdsSeteoF, sSQLAct, true);
						
						iCont++;			
					}

					this.Cursor = Cursors.Default;
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);					
				this.Cursor = Cursors.Default;	
				return;										
			}
			#endregion Sube archivo
				
			#region Cruce de cobros
			this.Cursor = Cursors.WaitCursor;

			DateTime dtFecha = DateTime.Today;

			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
			{
				oConexion.Open();

				SqlTransaction oTransaction = null;
				SqlCommand oCmdActualiza = oConexion.CreateCommand();
				oCmdActualiza.Connection = oConexion;
				oCmdActualiza.CommandType = CommandType.Text;
				oCmdActualiza.CommandTimeout = 120;
						
				try
				{							
					oTransaction = oConexion.BeginTransaction();
					oCmdActualiza.Transaction = oTransaction;
					
					SqlDataAdapter oDACM = new SqlDataAdapter(oCmdActualiza);
					DataSet oDSCM = new DataSet();						
					oCmdActualiza.CommandText = "Select idHistorial_Cash_Management, Ruc, Fecha, Valor, NoReferencia From Historial_Cash_Management Where Estado = 0 Order By Ruc";
					oDACM.Fill(oDSCM);
					
					foreach(DataRow oDR in oDSCM.Tables[0].Rows)
					{
						#region Valida si existe operacion de credito
						int iContOp = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From Compra Where idTipoFactura = 1 And Borrar = 0 And Estado = 10 And idFormaPago = 9 And idEntidadFinanciera <> 2 And Ruc = '{0}'", oDR["Ruc"].ToString().Trim()));						
						string sOperacion = "";
	
						if (iContOp == 1)
						{
							oCmdActualiza.CommandText = 
								string.Format("Select idCompra From Compra Where idTipoFactura = 1 And Borrar = 0 And Estado = 10 And idFormaPago = 9 And idEntidadFinanciera <> 2 And Ruc = '{0}'", oDR["Ruc"].ToString().Trim());
							idCompra = (int)oCmdActualiza.ExecuteScalar();

							oCmdActualiza.CommandText = string.Format("Select NumeroDeOperacion From Compra Where idTipoFactura = 1 And Borrar = 0 And Estado = 10 And idFormaPago = 9 And idEntidadFinanciera <> 2 And Ruc = '{0}'", oDR["Ruc"].ToString().Trim());
							sOperacion = oCmdActualiza.ExecuteScalar().ToString();
						}

						if (iContOp > 0)
						{

						}
						#endregion Valida si existe operacion de credito

						#region Genera Abono 
						oCmdActualiza.CommandText = "Exec NumeracionLocales 24, 57, 0";
						string sNumero = oCmdActualiza.ExecuteScalar().ToString(); 
							
						string sSQLPago = string.Format("Exec GuardaPagoCobranza {0}, {1}, {2}, 0, '{3}', '{4}', {5}", 
							idCompra, 57, decimal.Parse(oDR["Valor"].ToString()), DateTime.Parse(oDR["Fecha"].ToString()).ToString("yyyyMMdd"), sNumero, false);
						oCmdActualiza.CommandText = sSQLPago;
						idAnticipo = (int)oCmdActualiza.ExecuteScalar(); 
							
						string sSQLCobro = string.Format("Exec GuardaDetalleFormaPago {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', '{8}', '{9}', '{10}', '{11}', {12}, '{13}', '{14}', {15}, {16}, {17}, 0", 
							0, idAnticipo, 7, 13, decimal.Parse(oDR["Valor"].ToString()), 7, 0, oDR["NoReferencia"].ToString(),
							DateTime.Today.ToString("yyyyMMdd"), "", "", 12, "", "", 0, 0, 0);
						oCmdActualiza.CommandText = sSQLCobro;
						oCmdActualiza.ExecuteNonQuery(); 		
						#endregion Genera Abono 

						#region Variables
						decimal Cobro = dValor;
						decimal Gasto = 0.00m;
						decimal IntMora = 0.00m;
						decimal IntFin = 0.00m;
						decimal Capital = 0.00m;
						decimal Abono = 0.00m;
						decimal Saldo = dValor;
						decimal NSaldo = 0.00m;
						string sSQL = "";
						decimal AbGastos = 0.00m;
						decimal AbIntMora = 0.00m;
						decimal AbInteres = 0.00m;
						decimal AbCapital = 0.00m;
						int Estado = 0;
						string sConcepto = "";						
						#endregion Variables

						#region Calculo de Cobranza
						SqlDataAdapter oDACuotas = new SqlDataAdapter(oCmdActualiza);
						DataSet oDSCuotas = new DataSet();
						oCmdActualiza.CommandText = string.Format("Exec Cobranzas {0}, '{1}'", idCompra, dtFecha.ToString("yyyyMMdd"));
						oDACuotas.Fill(oDSCuotas);
						foreach(DataRow oDRCuotas in oDSCuotas.Tables[0].Rows)
						{
							if (Saldo > 0)
							{
								IntMora = 0.00m;
								IntFin = 0.00m;
								Capital = 0.00m;
								AbGastos = 0.00m;
								AbIntMora = 0.00m;
								AbInteres = 0.00m;
								AbCapital = 0.00m;
								Estado = 0;
								
								IntMora = Convert.ToDecimal(oDRCuotas["Mora"]);
								IntFin = Convert.ToDecimal(oDRCuotas["Interes"]);
								Capital = Convert.ToDecimal(oDRCuotas["Capital"]);									
								Abono = 0.00m;
								
								#region Interés por mora
								if (Math.Round(IntMora, 2) > 0.00m && Math.Round(Saldo, 2) > 0.00m && Math.Round(Saldo, 2) >= Math.Round(IntMora, 2))
								{
									Abono = Abono + IntMora;
									AbIntMora = AbIntMora + IntMora;
									Saldo = Saldo - IntMora;
								}
								else if (Math.Round(Saldo, 2) > 0.00m && Math.Round(IntMora, 2) > 0.00m && Math.Round(IntMora, 2) >= Math.Round(Saldo, 2))
								{
									IntMora = IntMora - Saldo;
									Abono = Abono + Saldo;
									AbIntMora = AbIntMora + Saldo;
									Saldo = 0.00m;
								}
								#endregion Interés por mora
						
								#region Interés por financiamiento
								if (Math.Round(IntFin, 2) > 0.00m && Math.Round(Saldo, 2) > 0.00m && Math.Round(Saldo, 2) >= Math.Round(IntFin, 2))
								{
									Abono = Abono + IntFin;
									AbInteres = AbInteres + IntFin;
									Saldo = Saldo - IntFin;
								}
								else if (Math.Round(Saldo, 2) > 0.00m && Math.Round(IntFin, 2) > 0.00m && Math.Round(IntFin, 2) >= Math.Round(Saldo, 2))
								{
									IntFin = IntFin - Saldo;
									Abono = Abono + Saldo;
									AbInteres = AbInteres + Saldo;
									Saldo = 0.00m;
								}
								#endregion Interés por financiamiento

								#region Capital
								if (Math.Round(Capital, 2) > 0.00m && Math.Round(Saldo, 2) > 0.00m && Math.Round(Saldo, 2) >= Math.Round(Capital, 2))
								{
									Abono = Abono + Capital;
									AbCapital = AbCapital + Capital;
									Saldo = Saldo - Capital;
								}
								else if (Math.Round(Saldo, 2) > 0.00m && Math.Round(Capital, 2) > 0.00m && Math.Round(Capital, 2) >= Math.Round(Saldo, 2))
								{
									Capital = Capital - Saldo;
									Abono = Abono + Saldo;
									AbCapital = AbCapital + Saldo;
									Saldo = 0.00m;								
								}
								#endregion Capital

								#region Actualiza tabla Cre_TablaDeAmortización
								if (Abono > 0.00m)
								{										
									decimal dSaldoPrb = 0.00m;
									dSaldoPrb = Convert.ToDecimal(oDRCuotas["Saldo"]);
									NSaldo = Convert.ToDecimal(oDRCuotas["Saldo"]) - Abono;

									string sSQLAC = string.Format("Exec ActualizaCobranza {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, '{8}', {9}, {10}, {11}, {12}, {13}", 
										(int)oDRCuotas["idCre_TablaDeAmortizacion"], Gasto, IntMora, IntFin, Capital, Abono, Math.Round(NSaldo, 2), (int)oDRCuotas["Retraso"], dtFecha.ToString("yyyyMMdd"), idCompra, AbInteres, AbCapital, AbGastos, AbIntMora);
									oCmdActualiza.CommandText = sSQLAC;
									oCmdActualiza.ExecuteNonQuery();
							
									sConcepto = "";
									if (Math.Round(NSaldo, 2) == 0.00m) 
									{
										Estado = 2;
										sConcepto = "CANCELACIÓN CUOTA CREDIPOINT NÚMERO " + oDRCuotas["NumeroCuota"].ToString();
									}
									else if (Math.Round(NSaldo, 2) > 0.00m) 
									{
										Estado = 3;
										sConcepto = "ABONO CUOTA CREDIPOINT NÚMERO " + oDRCuotas["NumeroCuota"].ToString();
									}
							
									string sSQLTipoPago = string.Format("Select Pago From Cre_TablaDeAmortizacion Where idCre_TablaDeAmortizacion = {0}", (int)oDRCuotas["idCre_TablaDeAmortizacion"]);
									oCmdActualiza.CommandText = sSQLTipoPago;
									int iPago = (int)oCmdActualiza.ExecuteScalar();

									decimal dDebitado = AbIntMora + AbInteres + AbCapital;
									string sSQLPCC = string.Format("Exec GuardaPagosCuotasCredito {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, '{9}', '{10}', {11}, {12}, 1, '{13}'",
										idCompra, (int)oDRCuotas["idCre_TablaDeAmortizacion"],
										AbGastos, AbIntMora, AbInteres, AbCapital, Abono, Math.Round(NSaldo, 2), Estado, sConcepto, 
										dtFecha.ToString("yyyyMMdd"), iPago, dDebitado, sOperacion);
									oCmdActualiza.CommandText = sSQLPCC;
									oCmdActualiza.ExecuteNonQuery();
								}
								#endregion Actualiza tabla Cre_TablaDeAmortización
							}
						}	
						#endregion Calculo de Cobranza
								
						#region Asiento - Saldos
						string sSQLAsiento = string.Format("Exec AsientoCobranzaCrediPoint {0}, 0, {1}", idAnticipo, idCompra);
						oCmdActualiza.CommandText = sSQLAsiento;
						oCmdActualiza.ExecuteNonQuery();
							
						string sSQLSaldo = string.Format("Exec ActualizaSaldoFacturas {0}, 24, 9 Exec ActualizaSaldoFacturas {0}, 1, 9", idAnticipo, idCompra);
						oCmdActualiza.CommandText = sSQLSaldo;
						oCmdActualiza.ExecuteNonQuery();
						#endregion Asiento - Saldos
							
						#region Actualiza Compromisos de Pago - Estado de Cuenta 
						string sSQLActCompromisoPago = string.Format("Exec CPActualizaCompromisoPago {0}, '{1}', {2}", 
							idCompra, dtFecha.ToString("yyyyMMdd"), decimal.Parse(oDR["Valor"].ToString()));
						oCmdActualiza.CommandText = sSQLActCompromisoPago;
						oCmdActualiza.ExecuteNonQuery();								

						string sSQLAEC = string.Format("Exec Cre_ActualizaEstadoDeCuentaDesdePago {0}", idCompra);
						oCmdActualiza.CommandText = sSQLAEC;
						oCmdActualiza.ExecuteNonQuery();
						#endregion Actualiza Compromisos de Pago - Estado de Cuenta

						#region Actualiza Valor Cobrado Gestiones de Credito_Cobranza
//						string sSQLActCompromisoCredito_Cobranza= string.Format("Exec Cbo_ActualizaCompromisoPago {0}, '{1}', {2}", 
//							idCompra, dtFecha.ToString("yyyyMMdd"), decimal.Parse(oDR["Valor"].ToString()));
//						oCmdActualiza.CommandText = sSQLActCompromisoCredito_Cobranza;
//						oCmdActualiza.ExecuteNonQuery();
						#endregion Actualiza Valor Cobrado Gestiones de Credito_Cobranza
					}
					oTransaction.Commit();	

					MessageBox.Show(string.Format("Importación Finalizada Correctamente, {0} Lineas Afectadas", iCont), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

					this.Cursor = Cursors.Default;	
				}
				catch (Exception ex)
				{
					oTransaction.Rollback();

					MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);

					oConexion.Close();					
				}
				finally
				{
					oConexion.Close();
				}
			}			
			#endregion Cruce de cobros			
		}
	}
}
