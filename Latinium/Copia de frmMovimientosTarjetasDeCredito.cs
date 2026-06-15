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
	/// Summary description for frmMovimientosTarjetasDeCredito.
	/// </summary>
	public class frmMovimientosTarjetasDeCredito : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridEntrada;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbBanco;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Button btnImportar;
		private System.Windows.Forms.OpenFileDialog openFileDialog1;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmMovimientosTarjetasDeCredito()
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
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			this.uGridEntrada = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cmbBanco = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.btnImportar = new System.Windows.Forms.Button();
			this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.uGridEntrada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridEntrada
			// 
			this.uGridEntrada.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridEntrada.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridEntrada.DisplayLayout.Appearance = appearance1;
			this.uGridEntrada.DisplayLayout.AutoFitColumns = true;
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridEntrada.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.uGridEntrada.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridEntrada.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.uGridEntrada.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 8.25F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridEntrada.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.uGridEntrada.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance5.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance5.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridEntrada.DisplayLayout.Override.RowAlternateAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridEntrada.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridEntrada.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.uGridEntrada.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridEntrada.Location = new System.Drawing.Point(8, 40);
			this.uGridEntrada.Name = "uGridEntrada";
			this.uGridEntrada.Size = new System.Drawing.Size(1088, 280);
			this.uGridEntrada.TabIndex = 656;
			this.uGridEntrada.UpdateMode = Infragistics.Win.UltraWinGrid.UpdateMode.OnCellChange;
			// 
			// cmbBanco
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBanco.Appearance = appearance8;
			this.cmbBanco.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBanco.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbBanco.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBanco.DisplayMember = "Nombre";
			this.cmbBanco.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBanco.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBanco.Location = new System.Drawing.Point(80, 8);
			this.cmbBanco.Name = "cmbBanco";
			this.cmbBanco.Size = new System.Drawing.Size(248, 21);
			this.cmbBanco.TabIndex = 658;
			this.cmbBanco.ValueMember = "idBanco";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(16, 8);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(34, 16);
			this.label2.TabIndex = 657;
			this.label2.Text = "Banco";
			// 
			// btnImportar
			// 
			this.btnImportar.CausesValidation = false;
			this.btnImportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImportar.Location = new System.Drawing.Point(344, 8);
			this.btnImportar.Name = "btnImportar";
			this.btnImportar.TabIndex = 659;
			this.btnImportar.Text = "Importar";
			this.btnImportar.Click += new System.EventHandler(this.btnImportar_Click);
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
			// frmMovimientosTarjetasDeCredito
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1104, 326);
			this.Controls.Add(this.btnImportar);
			this.Controls.Add(this.cmbBanco);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.uGridEntrada);
			this.Name = "frmMovimientosTarjetasDeCredito";
			this.Text = "Movimientos Tarjetas De Credito";
			((System.ComponentModel.ISupportInitialize)(this.uGridEntrada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnImportar_Click(object sender, System.EventArgs e)
		{
			if(openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && openFileDialog1.FileName.Length > 0) 
			{
				string path = this.openFileDialog1.FileName;
 
				if(File.Exists(path))
				{
					OleDbConnection oConn = new OleDbConnection();
 
					OleDbCommand oCmd = new OleDbCommand();
 
					OleDbDataAdapter oDa = new OleDbDataAdapter();
 
					DataSet oDs = new DataSet();

					oConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + "; Extended Properties= Excel 8.0;";
 
					oConn.Open();
 
					oCmd.CommandText = "SELECT * FROM [ESTADODECUENTA$]";
 
					oCmd.Connection = oConn;
 
					oDa.SelectCommand = oCmd;
 
					oDa.Fill(oDs, "EstadoDeCuenta");

					oConn.Close();
	
					this.uGridEntrada.DataSource = oDs;
					
					foreach (DataRow row in oDs.Tables[0].Rows)
					{
						#region Variables
						int idMovimiento = 0;
						int idBanco = 6; // this.cmbBanco.Value;
						DateTime dtFecha = DateTime.Today;
						string sRecap = "";
						string sTipo = "";
						string sCuenta = "";
						string sNumeroTarjeta = "";
						decimal Tarifa12 = 0.00m;
						decimal Tarifa0 = 0.00m;
						decimal IVA = 0.00m;
						decimal ICE = 0.00m;
						decimal NoComisionable = 0.00m;
						decimal ValorTransaccion = 0.00m;
						decimal ValorComision = 0.00m;
						decimal RetencionIVA = 0.00m;
						decimal RetencionFuente = 0.00m;
						decimal ValordePago = 0.00m;						
						string sNumeroRetencion = "";
						string sAutorizacion = "";
						DateTime dtFechaCad = DateTime.Today;
						#endregion Variables

						#region Parametros
						/*
							@idMovimientoTarjeta Int, @idBanco Int, @FechaPago Datetime, @Recap Varchar(20), @Tipo Varchar(10), @Cuenta Varchar(30), @NumeroTarjeta Varchar(30), 
							@Tarifa12 Decimal(18, 2), @Tarifa0 Decimal(18, 2), @IVA Decimal(18, 2), @ICE Decimal(18, 2), 
							@NoComisionable Decimal(18, 2), @ValorTransaccion Decimal(18, 2), @ValorComision Decimal(18, 2), 
							@RetencionIVA Decimal(18, 2), @RetencionFuente Decimal(18, 2), @ValordePago Decimal(18, 2), 
							@NumeroRetencion Varchar(20), @Autorizacion Varchar(40), @Fecha Datetime
						*/

						/*
						 Fecha de Pago	Numero de Recap	Tipo de Credito	Numero de Cuenta Numero de Tarjeta	
						 Consumo Tarifa 12%	Consumo Tarifa 0%	Valor IVA	Valor ICE	
						 Valor No Comisionable	Valor Transaccion	Valor Comision	
						 Retencion IVA	Retencion Fuente	Valor de Pago	
						 Comprob de Retencion	Numero SRI autorizacion	Fecha vigencia autorizacion
						*/
						#endregion Parametros

						#region Asigna Valores
						if (row["Fecha de Pago"] != System.DBNull.Value) dtFecha = DateTime.Parse(row["Fecha de Pago"].ToString());
						if (row["Numero de Recap"] != System.DBNull.Value) sRecap = row["Numero de Recap"].ToString();
						if (row["Tipo de Credito"] != System.DBNull.Value) sTipo = row["Tipo de Credito"].ToString();
						if (row["Numero de Cuenta"] != System.DBNull.Value) sCuenta = row["Numero de Cuenta"].ToString();
						if (row["Numero de Tarjeta"] != System.DBNull.Value) sNumeroTarjeta = row["Numero de Tarjeta"].ToString();

						if (row["Consumo Tarifa 12%"] != System.DBNull.Value) Tarifa12 = decimal.Parse(row["Consumo Tarifa 12%"].ToString());
						if (row["Consumo Tarifa 0%"] != System.DBNull.Value) Tarifa0 = decimal.Parse(row["Consumo Tarifa 0%"].ToString());
						if (row["Valor IVA"] != System.DBNull.Value) IVA = decimal.Parse(row["Valor IVA"].ToString());
						if (row["Valor ICE"] != System.DBNull.Value) ICE = decimal.Parse(row["Valor ICE"].ToString());

						if (row["Valor No Comisionable"] != System.DBNull.Value) NoComisionable = decimal.Parse(row["Valor No Comisionable"].ToString());
						if (row["Valor Transaccion"] != System.DBNull.Value) ValorTransaccion = decimal.Parse(row["Valor Transaccion"].ToString());
						if (row["Valor Comision"] != System.DBNull.Value) ValorComision = decimal.Parse(row["Valor Comision"].ToString());

						if (row["Retencion IVA"] != System.DBNull.Value) RetencionIVA = decimal.Parse(row["Retencion IVA"].ToString());
						if (row["Retencion Fuente"] != System.DBNull.Value) RetencionFuente = decimal.Parse(row["Retencion Fuente"].ToString());
						if (row["Valor de Pago"] != System.DBNull.Value) ValordePago = decimal.Parse(row["Valor de Pago"].ToString());

						if (row["Comprob de Retencion"] != System.DBNull.Value) sNumeroRetencion = row["Comprob de Retencion"].ToString();
						if (row["Numero SRI autorizacion"] != System.DBNull.Value) sAutorizacion = row["Numero SRI autorizacion"].ToString();
						if (row["Fecha vigencia autorizacion"] != System.DBNull.Value) dtFechaCad = DateTime.Parse(row["Fecha vigencia autorizacion"].ToString());
						#endregion Asigna Valores

						string sSQLValida = string.Format("Exec ValidaVouchersTarjetasdeCredito '{0}', '{1}', {2}, {3}", sRecap, sNumeroTarjeta, idBanco, decimal.Parse(row["Valor IVA"].ToString()));
						SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQLValida, true);
						dr.Read();
						int idCompra = dr.GetInt32(0);
						int iOrigen = dr.GetInt32(1);
						dr.Close();

						int iEstado = 0;
						if (idCompra > 0) iEstado = 1;

						string sSQLInserta = string.Format("Exec InsertaMovimientoTarjetas {0}, {1}, '{2}', '{3}', '{4}', '{5}', '{6}', {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, '{17}', '{18}', '{19}', {20}, {21}, {22}", 
							idMovimiento, idBanco, dtFecha.ToString("yyyyMMdd"), sRecap.Substring(2, 4), sTipo, sCuenta, sNumeroTarjeta.Substring(0, 6),
							Tarifa12, Tarifa0, IVA, ICE,
							NoComisionable, ValorTransaccion, ValorComision,
							RetencionIVA, RetencionFuente, ValordePago, 
							sNumeroRetencion, sAutorizacion, dtFechaCad.ToString("yyyyMMdd"), iEstado, idCompra, iOrigen);
						Funcion.EjecutaSQL(cdsSeteoF, sSQLInserta);
					}
					
					this.uGridEntrada.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select * From MovimientosTarjetas");
				}
			}
		}
		

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}
	}
}

