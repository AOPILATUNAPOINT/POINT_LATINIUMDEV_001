using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.IO;


namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmLiquidacionBancoAustro.
	/// </summary>
	public class frmLiquidacionBancoAustro : System.Windows.Forms.Form
	{
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnContabilizar;
		private System.Windows.Forms.Button btnVerAsiento;
		private System.Windows.Forms.Button btnExportar;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Label lblContador;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidAsiento;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo txtFecha;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridLotes;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtLote;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.OpenFileDialog openFileDialog1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		public bool bCarga = false;

		public frmLiquidacionBancoAustro()
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
		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}
		public bool ActualizaNumeroOperacion(string stArchivo) 
		{ 			 
			if(File.Exists(stArchivo))
			{
				OleDbConnection oConn = new OleDbConnection(); 
				OleDbCommand oCmd = new OleDbCommand(); 
				OleDbDataAdapter oDa = new OleDbDataAdapter(); 
				DataSet oDs = new DataSet();
				oConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + stArchivo + "; Extended Properties= Excel 8.0;"; 
				oConn.Open(); 
				oCmd.CommandText = "SELECT * FROM [Lote$]"; 
				oCmd.Connection = oConn; 
				oDa.SelectCommand = oCmd; 
				oDa.Fill(oDs, "Lote");
 
				foreach (DataRow row in oDs.Tables[0].Rows)
				{					
					string sNumeroDeOperacion = "";
					string sCedula = "";
					decimal dCapital = 0.00m;
					decimal dDescuento= 0.00m;
					decimal dSolca= 0.00m;
					
					if (row["Identificacion"] != System.DBNull.Value) sCedula = row["Identificacion"].ToString();
					if (row["Operacion"] != System.DBNull.Value) sNumeroDeOperacion = row["Operacion"].ToString();
					if (row["Monto Saldo"] != System.DBNull.Value) dCapital = decimal.Parse(row["Monto Saldo"].ToString());
					if (row["Descuento"] != System.DBNull.Value) dDescuento = decimal.Parse(row["Descuento"].ToString());
					if (row["Solca"] != System.DBNull.Value) dSolca = decimal.Parse(row["Solca"].ToString());

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLotes.Rows.All)
					{
						if (sCedula.Trim().Length == 9) sCedula = "0" + sCedula.Trim();
						
						if (sCedula.Trim() == dr.Cells["Ruc"].Value.ToString().Trim())
						{
							dCapital=dCapital-dDescuento-dSolca;

							if (dCapital == Convert.ToDecimal(dr.Cells["CapitalVendido"].Value))
							{
								dr.Cells["Operacion_GAF"].Value = sNumeroDeOperacion.Trim();
								dr.Cells["Comision"].Value = dDescuento;

								break;
							}
						}
					}
				}

				return true; 
			}
			else
			{
				return false; 
			}
		}
		

		#region Código generado por el Diseñador de Windows Forms
		/// <summary>
		/// Método necesario para admitir el Diseñador. No se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmLiquidacionBancoAustro));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsientoLoteSolidario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Num_Operacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Operacion_GAF");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Monto");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Plazo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuotas_Locales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuotas_Vendidas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CapitalVendido");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Contribucion_Solca");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 15, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 15, true);
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Contribucion_Solca", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Contribucion_Solca", 14, true);
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Comision", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Comision", 13, true);
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CapitalVendido", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CapitalVendido", 12, true);
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsientoLoteSolidario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Num_Operacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Operacion_GAF");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Monto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Plazo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaVenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuotas_Locales");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuotas_Vendidas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CapitalVendido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Comision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Contribucion_Solca");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.btnVer = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnContabilizar = new System.Windows.Forms.Button();
			this.btnVerAsiento = new System.Windows.Forms.Button();
			this.btnExportar = new System.Windows.Forms.Button();
			this.lblEstado = new System.Windows.Forms.Label();
			this.lblContador = new System.Windows.Forms.Label();
			this.txtidAsiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.txtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.uGridLotes = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtLote = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.button1 = new System.Windows.Forms.Button();
			this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
			((System.ComponentModel.ISupportInitialize)(this.txtidAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLotes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLote)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// groupBox2
			// 
			this.groupBox2.Location = new System.Drawing.Point(8, 40);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1040, 8);
			this.groupBox2.TabIndex = 775;
			this.groupBox2.TabStop = false;
			this.groupBox2.Enter += new System.EventHandler(this.groupBox2_Enter);
			// 
			// btnVer
			// 
			this.btnVer.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnVer.CausesValidation = false;
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnVer.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(464, 8);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(80, 23);
			this.btnVer.TabIndex = 774;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(752, 8);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(88, 23);
			this.btnImprimir.TabIndex = 773;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnContabilizar
			// 
			this.btnContabilizar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnContabilizar.CausesValidation = false;
			this.btnContabilizar.Enabled = false;
			this.btnContabilizar.Image = ((System.Drawing.Image)(resources.GetObject("btnContabilizar.Image")));
			this.btnContabilizar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnContabilizar.Location = new System.Drawing.Point(944, 56);
			this.btnContabilizar.Name = "btnContabilizar";
			this.btnContabilizar.Size = new System.Drawing.Size(88, 23);
			this.btnContabilizar.TabIndex = 772;
			this.btnContabilizar.Text = "Contabilizar";
			this.btnContabilizar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnContabilizar.Click += new System.EventHandler(this.btnContabilizar_Click);
			// 
			// btnVerAsiento
			// 
			this.btnVerAsiento.CausesValidation = false;
			this.btnVerAsiento.Enabled = false;
			this.btnVerAsiento.Image = ((System.Drawing.Image)(resources.GetObject("btnVerAsiento.Image")));
			this.btnVerAsiento.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVerAsiento.Location = new System.Drawing.Point(848, 8);
			this.btnVerAsiento.Name = "btnVerAsiento";
			this.btnVerAsiento.Size = new System.Drawing.Size(88, 23);
			this.btnVerAsiento.TabIndex = 771;
			this.btnVerAsiento.Text = "Ver Asiento";
			this.btnVerAsiento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVerAsiento.Click += new System.EventHandler(this.btnVerAsiento_Click);
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(944, 8);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(88, 23);
			this.btnExportar.TabIndex = 770;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(24, 64);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 20);
			this.lblEstado.TabIndex = 769;
			this.lblEstado.Click += new System.EventHandler(this.lblEstado_Click);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.BackColor = System.Drawing.Color.Transparent;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(488, 56);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 20);
			this.lblContador.TabIndex = 768;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblContador.Click += new System.EventHandler(this.lblContador_Click);
			// 
			// txtidAsiento
			// 
			this.txtidAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtidAsiento.Appearance = appearance1;
			this.txtidAsiento.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtidAsiento.Enabled = false;
			this.txtidAsiento.FormatString = "";
			this.txtidAsiento.Location = new System.Drawing.Point(616, 8);
			this.txtidAsiento.Name = "txtidAsiento";
			this.txtidAsiento.PromptChar = ' ';
			this.txtidAsiento.Size = new System.Drawing.Size(8, 22);
			this.txtidAsiento.TabIndex = 767;
			this.txtidAsiento.Visible = false;
			this.txtidAsiento.ValueChanged += new System.EventHandler(this.txtidAsiento_ValueChanged);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(8, 80);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1032, 8);
			this.groupBox1.TabIndex = 766;
			this.groupBox1.TabStop = false;
			this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(216, 8);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(97, 17);
			this.label2.TabIndex = 765;
			this.label2.Text = "Fecha De Compra:";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.label2.Click += new System.EventHandler(this.label2_Click);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(16, 11);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(73, 17);
			this.label3.TabIndex = 762;
			this.label3.Text = "Lote Número:";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label3.Click += new System.EventHandler(this.label3_Click);
			// 
			// txtFecha
			// 
			this.txtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.txtFecha.DateButtons.Add(dateButton1);
			this.txtFecha.Format = "dd/MM/yyyy";
			this.txtFecha.Location = new System.Drawing.Point(336, 8);
			this.txtFecha.Name = "txtFecha";
			this.txtFecha.NonAutoSizeHeight = 23;
			this.txtFecha.Size = new System.Drawing.Size(112, 21);
			this.txtFecha.SpinButtonsVisible = true;
			this.txtFecha.TabIndex = 764;
			this.txtFecha.Value = ((object)(resources.GetObject("txtFecha.Value")));
			this.txtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.txtFecha_BeforeDropDown);
			// 
			// uGridLotes
			// 
			this.uGridLotes.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridLotes.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridLotes.DataSource = this.ultraDataSource1;
			appearance2.BackColor = System.Drawing.Color.White;
			this.uGridLotes.DisplayLayout.Appearance = appearance2;
			this.uGridLotes.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 77;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 64;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.Caption = "Factura N.";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 127;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.Caption = "Operación";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 82;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 72;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.Caption = "Cedula";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 68;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 245;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance3;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.Caption = "Liquido";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 89;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 92;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.Caption = "Venta";
			ultraGridColumn10.Header.VisiblePosition = 12;
			ultraGridColumn10.Width = 67;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 9;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 74;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.Caption = "C Vendidas";
			ultraGridColumn12.Header.VisiblePosition = 10;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 72;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance4;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.Caption = "Capital";
			ultraGridColumn13.Header.VisiblePosition = 11;
			ultraGridColumn13.Width = 84;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance5;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Width = 68;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance6;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.Caption = "Impuesto";
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Width = 77;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance7;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.Caption = "Acreditado";
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Width = 129;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance8;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance9;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance10;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance11;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridLotes.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance12.ForeColor = System.Drawing.Color.Black;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLotes.DisplayLayout.GroupByBox.Appearance = appearance12;
			appearance13.ForeColor = System.Drawing.Color.Black;
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLotes.DisplayLayout.GroupByBox.BandLabelAppearance = appearance13;
			this.uGridLotes.DisplayLayout.GroupByBox.Hidden = true;
			this.uGridLotes.DisplayLayout.GroupByBox.Prompt = " ";
			appearance14.BackColor = System.Drawing.SystemColors.Control;
			appearance14.BackColor2 = System.Drawing.SystemColors.Control;
			appearance14.BackColorDisabled = System.Drawing.SystemColors.Control;
			appearance14.BackColorDisabled2 = System.Drawing.SystemColors.Control;
			appearance14.ForeColor = System.Drawing.Color.Black;
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLotes.DisplayLayout.GroupByBox.PromptAppearance = appearance14;
			appearance15.ForeColor = System.Drawing.Color.Black;
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLotes.DisplayLayout.Override.ActiveRowAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.Transparent;
			this.uGridLotes.DisplayLayout.Override.CardAreaAppearance = appearance16;
			this.uGridLotes.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.AllRowsInBand;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance17.FontData.BoldAsString = "True";
			appearance17.FontData.Name = "Arial";
			appearance17.FontData.SizeInPoints = 8F;
			appearance17.ForeColor = System.Drawing.Color.White;
			appearance17.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridLotes.DisplayLayout.Override.HeaderAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLotes.DisplayLayout.Override.RowSelectorAppearance = appearance18;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLotes.DisplayLayout.Override.SelectedRowAppearance = appearance19;
			this.uGridLotes.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridLotes.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridLotes.Location = new System.Drawing.Point(8, 88);
			this.uGridLotes.Name = "uGridLotes";
			this.uGridLotes.Size = new System.Drawing.Size(1040, 288);
			this.uGridLotes.TabIndex = 763;
			this.uGridLotes.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridLotes_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(System.DateTime);
			ultraDataColumn5.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn8.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(System.DateTime);
			ultraDataColumn10.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn11.DefaultValue = 0;
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn12.DefaultValue = 0;
			ultraDataColumn13.DataType = typeof(System.Decimal);
			ultraDataColumn13.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn14.DataType = typeof(System.Decimal);
			ultraDataColumn14.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn15.DataType = typeof(System.Decimal);
			ultraDataColumn15.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn16.DataType = typeof(System.Decimal);
			ultraDataColumn16.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16});
			// 
			// txtLote
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtLote.Appearance = appearance20;
			this.txtLote.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtLote.FormatString = "";
			this.txtLote.Location = new System.Drawing.Point(112, 8);
			this.txtLote.MinValue = -1;
			this.txtLote.Name = "txtLote";
			this.txtLote.PromptChar = ' ';
			this.txtLote.Size = new System.Drawing.Size(88, 22);
			this.txtLote.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtLote.TabIndex = 761;
			this.txtLote.ValueChanged += new System.EventHandler(this.txtLote_ValueChanged);
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
			// button1
			// 
			this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button1.CausesValidation = false;
			this.button1.Enabled = false;
			this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
			this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button1.Location = new System.Drawing.Point(484, 180);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(88, 23);
			this.button1.TabIndex = 776;
			this.button1.Text = "Contabilizar";
			this.button1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// frmLiquidacionBancoAustro
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.ClientSize = new System.Drawing.Size(1056, 382);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnContabilizar);
			this.Controls.Add(this.btnVerAsiento);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.txtidAsiento);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.txtFecha);
			this.Controls.Add(this.uGridLotes);
			this.Controls.Add(this.txtLote);
			this.Controls.Add(this.button1);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.Name = "frmLiquidacionBancoAustro";
			this.Text = "Liquidacion Banco Austro";
			this.Load += new System.EventHandler(this.frmLiquidacionBancoAustro_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtidAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLotes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLote)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
				FuncionesProcedimientos.ExportaExcel(this.uGridLotes);
		}

		private void lblEstado_Click(object sender, System.EventArgs e)
		{
		
		}

		private void lblContador_Click(object sender, System.EventArgs e)
		{
		
		}

		private void txtidAsiento_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void groupBox1_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void txtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void label2_Click(object sender, System.EventArgs e)
		{
		
		}

		private void uGridLotes_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void groupBox2_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			if (this.txtLote.Value == System.DBNull.Value)
			{
				MessageBox.Show("Escriba el Numero de Lote a Consultar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtLote.Focus();
				return;
			}

			this.lblEstado.Text = "";
			this.lblContador.Text = "";

			this.txtFecha.Enabled = false;
			//this.btnContabilizar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnVerAsiento.Enabled = false;

			string sSQL = string.Format("Exec Cre_LiquidacionCarteraAustro {0}, '{1}'", 
				(int)this.txtLote.Value, Convert.ToDateTime(this.txtFecha.Value).ToString("yyyyMMdd"));
			FuncionesProcedimientos.dsGeneralModulo(sSQL, this.uGridLotes);
			
			this.lblContador.Text = this.uGridLotes.Rows.Count + " REGISTROS ENCONTRADOS";
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			if (this.txtLote.Value == System.DBNull.Value)
			{
				MessageBox.Show("Escriba el Numero de Lote", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtLote.Focus();
				return;
			}

			if (this.uGridLotes.Rows.Count == 0)
			{
				MessageBox.Show(string.Format("No hay Documentos asignados en el Lote {0}", (int)this.txtLote.Value), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				this.txtLote.Focus();
				return;
			}

			if ((int)this.txtidAsiento.Value == 0)
			{
				MessageBox.Show(string.Format("No se ha Contabilizado el Lote {0}", (int)this.txtLote.Value), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				this.txtLote.Focus();
				return;
			}
			
			#region Impresion
			string strFiltro = "{Asiento.idAsiento} = " + this.txtidAsiento.Value.ToString();
				
			Reporte miRepor = new Reporte("Asientos.rpt", strFiltro);
			miRepor.MdiParent = this.MdiParent;
			miRepor.Titulo("Asiento");			
			miRepor.Show();	
			#endregion Impresion			
		}

		private void txtLote_ValueChanged(object sender, System.EventArgs e)
		{
			this.txtidAsiento.Value =
				Funcion.iEscalarSQL(cdsSeteoF, string.Format("Declare @idAsiento Int = 0 If Exists(Select ISNULL(idAsiento, 0) From Asiento Where Numero = '{0}' And idOrigenAsiento = 103 And idTipoAsiento = 5 And Borrar = 0) Select @idAsiento = ISNULL(idAsiento, 0) From Asiento Where Numero = '{0}' And idTipoAsiento = 5 And Borrar = 0 Select @idAsiento", this.txtLote.Value.ToString()));

			if ((int)this.txtidAsiento.Value > 0)
			{				
				this.lblEstado.Text = "CONTABILIZADO";
				this.txtFecha.Enabled = false;
				this.btnImprimir.Enabled = true;
				this.btnVerAsiento.Enabled = true;

				this.txtFecha.Value = Funcion.dtEscalarSQL(cdsSeteoF, string.Format("Select Fecha From Asiento Where idAsiento = {0}", (int)this.txtidAsiento.Value));
			}
			else 
			{
				this.lblEstado.Text = "CONTABILIZACIÓN PENDIENTE";				
				this.txtFecha.Enabled = true;
				this.btnContabilizar.Enabled = true;				
			}
		}

		private void btnContabilizar_Click(object sender, System.EventArgs e)
		{			
			#region Validacion
			if (this.txtLote.Value == System.DBNull.Value)
			{
				MessageBox.Show("Escriba el Numero de Lote", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);			
				this.txtLote.Focus();
				return;
			}

			if (this.uGridLotes.Rows.Count == 0)
			{
				MessageBox.Show("No hay Documentos con este Lote", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				this.txtLote.Focus();
				return;
			}

			if (!Validacion.vbFechaEnDocumentos(this.txtFecha, "Ingrese la fecha del documento", true, "Bancos", cdsSeteoF)) return;
			
			DateTime dtUltLiq = Funcion.dtEscalarSQL(cdsSeteoF, "Select MAX(ISNULL(FechaVenta, 0)) From Compra Where idTipoFactura = 1 And Borrar = 0 And Estado In (10, 11) And idFormaPago = 9 And idEntidadFinanciera = 15 And CapturaLote = 1");

			if (Convert.ToDateTime(this.txtFecha.Value) < dtUltLiq)
			{				
				MessageBox.Show(string.Format("La Fecha de Liquidación no puede ser anterior a '{0}'", dtUltLiq.ToString("yyyy/MM/dd")), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtFecha.Value = dtUltLiq;
				this.txtFecha.Focus();			
				return;
			}
			#endregion Validacion
			
			try
			{		
				if (DialogResult.Yes == MessageBox.Show(string.Format("Esta seguro de Contabilizar el Lote Numero {0}", (int)this.txtLote.Value), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
				{
					OpenFileDialog sArchivo = new OpenFileDialog();

					sArchivo.Filter = "Microsoft Excel (*.xls)|*.xls";

					if (sArchivo.ShowDialog() == System.Windows.Forms.DialogResult.OK && sArchivo.FileName.Length > 0)
					{
						this.Cursor = Cursors.WaitCursor;
					
						#region Variables
						int iContVigente = 0;

						string	NUMERO_OPERACION	="";
						decimal	CAPITAL_VENDIDO	= 0;
						decimal COMISION =  0;	
						decimal CONTRIBUCION_SOLCA = 0;
						string	CEDULA ="";
						decimal	MONTO_ORIGINAL	= 0;

						#endregion Variables

						#region Sube informacion
						foreach (DataRow row in FuncionesProcedimientos.ImportaDeExcel(sArchivo.FileName, "SELECT * FROM [DATOS$]", "DATOS").Tables[0].Rows)
						{	
							NUMERO_OPERACION = "";
							CAPITAL_VENDIDO	= 0;
							COMISION = 0;	
							CONTRIBUCION_SOLCA = 0;
							CEDULA ="";
							MONTO_ORIGINAL = 0;

							if (row["Número Operación"] != System.DBNull.Value)  NUMERO_OPERACION = row["Número Operación"].ToString();
							if (row["Monto Saldo"] != System.DBNull.Value) CAPITAL_VENDIDO = decimal.Parse( row["Monto Saldo"].ToString()); 
							if (row["Fondo Garantía"] != System.DBNull.Value) COMISION = decimal.Parse( row["Fondo Garantía"].ToString()); 
							if (row["Solca"] != System.DBNull.Value) CONTRIBUCION_SOLCA = decimal.Parse( row["Solca"].ToString());
							if (row["Identificación "] != System.DBNull.Value) CEDULA = row["Identificación "].ToString();
							if (row["Monto Original"] != System.DBNull.Value) MONTO_ORIGINAL = decimal.Parse( row["Monto Original"].ToString());
						
							string sSQLAct = string.Format("Exec ActualizaValoresLiquidacionAustro '{0}', {1}, {2}, {3}, '{4}', {5},'{6}', {7}", 
								NUMERO_OPERACION,CAPITAL_VENDIDO,COMISION,CONTRIBUCION_SOLCA,Convert.ToDateTime(this.txtFecha.Value).ToString("yyyyMMdd"),(int)this.txtLote.Value,CEDULA,MONTO_ORIGINAL);						
							Funcion.EjecutaSQL(cdsSeteoF, sSQLAct, true);
						
							iContVigente++;			
						}			
						#endregion Sube informacion

						string sSQLALC = string.Format("Exec AsientoLiquidacionCarteraBancoAustro {0}, {1}, '{2}'", 
							(int)this.txtidAsiento.Value,(int)this.txtLote.Value,Convert.ToDateTime(this.txtFecha.Value).ToString("yyyyMMdd"));						
						int idAsiento = Funcion.iEscalarSQL(cdsSeteoF, sSQLALC, true);

						string sSQLCLTC = string.Format("Exec Cre_LiquidaTablasCarteraAustro {0}",(int)this.txtLote.Value);						
						Funcion.iEscalarSQL(cdsSeteoF, sSQLCLTC, true);


						MessageBox.Show(string.Format("Se contabilizaron {0} operaciones correctamente", iContVigente), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
		
						#region Impresion
						string strFiltro = "{Asiento.idAsiento} = " + idAsiento.ToString();				
						Reporte miRepor = new Reporte("Asientos.rpt", strFiltro);
						miRepor.MdiParent = this.MdiParent;
						miRepor.Titulo("Asiento");			
						miRepor.Show();	
						#endregion Impresion

						this.btnVer_Click(sender, e);

						this.Cursor = Cursors.Default;
					}	
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Demo", MessageBoxButtons.OK, MessageBoxIcon.Error);					
				this.Cursor = Cursors.Default;													
			}
		}

		private void btnVerAsiento_Click(object sender, System.EventArgs e)
		{
			if (this.txtLote.Value == System.DBNull.Value)
			{
				MessageBox.Show("Escriba el Numero de Lote", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);		
				this.txtLote.Focus();
				return;
			}

			if (this.uGridLotes.Rows.Count == 0)
			{
				MessageBox.Show(string.Format("No hay Documentos asignados en el Lote {0}", (int)this.txtLote.Value), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);	
				this.txtLote.Focus();
				return;
			}

			if ((int)this.txtidAsiento.Value == 0)
			{
				MessageBox.Show(string.Format("No se ha Contabilizado el Lote {0}", (int)this.txtLote.Value), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				this.txtLote.Focus();
				return;
			}
			
			Cursor = Cursors.WaitCursor;
			frmAsientos Asiento = new frmAsientos(((int)this.txtidAsiento.Value));
			Asiento.MdiParent = MdiParent;
			Asiento.Show();
			Cursor = Cursors.Default;			
		}

		private void label3_Click(object sender, System.EventArgs e)
		{
		
		}

		private void frmLiquidacionBancoAustro_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "08320201");

			if (!Funcion.Permiso("961", cdsSeteoF))
			{		
				MessageBox.Show("No tiene Acceso LIQUIDACION CARTERA AUSTRO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			
			bCarga = true;
			this.txtLote.Value = Funcion.iEscalarSQL(cdsSeteoF, "Cre_RetornaUltimoLote 15");
			this.txtLote.MaxValue = (int)this.txtLote.Value;
			
			this.txtFecha.CalendarInfo.MinDate = DateTime.Parse("25/08/2022");
			this.txtFecha.CalendarInfo.MaxDate = DateTime.Today;
			this.txtFecha.Value = DateTime.Today;

			bCarga = false;

			FuncionesProcedimientos.RetornaFechaServidor(this.txtFecha, cdsSeteoF, false);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}
	}
}
