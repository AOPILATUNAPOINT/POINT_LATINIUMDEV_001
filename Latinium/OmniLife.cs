using System;
using System.IO;
using System.Globalization;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using C1.C1Excel;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de OmniLife.
	/// </summary>
	public class OmniLife : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraButton btGanancia;
		private C1.Data.C1DataSet cdsCompra;
		private C1.C1Excel.C1XLBook c1XLBook1;
		private System.Windows.Forms.OpenFileDialog openFileDialog;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbFecha;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.Misc.UltraButton btImpuestos;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label label2;
		private C1.Data.C1DataSet cdsCuentaCorr;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCuentaCorr;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.Misc.UltraButton btInternacional;
		private C1.Data.C1DataSet cdsSeteo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNota;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.Misc.UltraButton btCodigos;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		private System.ComponentModel.IContainer components=null;

		public OmniLife()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CuentaCorriente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaCorriente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Agencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Automatico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaTipo");
			this.btGanancia = new Infragistics.Win.Misc.UltraButton();
			this.cdsCompra = new C1.Data.C1DataSet();
			this.c1XLBook1 = new C1.C1Excel.C1XLBook();
			this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
			this.cmbFecha = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.btImpuestos = new Infragistics.Win.Misc.UltraButton();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbCuentaCorr = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsCuentaCorr = new C1.Data.C1DataSet();
			this.label2 = new System.Windows.Forms.Label();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.btInternacional = new Infragistics.Win.Misc.UltraButton();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.txtNota = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.btCodigos = new Infragistics.Win.Misc.UltraButton();
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorr)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentaCorr)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNota)).BeginInit();
			this.SuspendLayout();
			// 
			// btGanancia
			// 
			this.btGanancia.Location = new System.Drawing.Point(128, 216);
			this.btGanancia.Name = "btGanancia";
			this.btGanancia.Size = new System.Drawing.Size(160, 23);
			this.btGanancia.TabIndex = 0;
			this.btGanancia.Text = "&Ganancias";
			this.btGanancia.Click += new System.EventHandler(this.btGanancia_Click);
			// 
			// cdsCompra
			// 
			this.cdsCompra.BindingContextCtrl = this;
			this.cdsCompra.DataLibrary = "LibFacturacion";
			this.cdsCompra.DataLibraryUrl = "";
			this.cdsCompra.DataSetDef = "dsCompra";
			this.cdsCompra.FillOnRequest = false;
			this.cdsCompra.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCompra.Name = "cdsCompra";
			this.cdsCompra.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCompra.SchemaDef = null;
			this.cdsCompra.BeforeFill += new C1.Data.FillEventHandler(this.cdsCompra_BeforeFill);
			// 
			// cmbFecha
			// 
			this.cmbFecha.DateTime = new System.DateTime(2006, 10, 15, 0, 0, 0, 0);
			this.cmbFecha.FormatString = "dd/MMM/yyyy";
			this.cmbFecha.Location = new System.Drawing.Point(152, 24);
			this.cmbFecha.Name = "cmbFecha";
			this.cmbFecha.Size = new System.Drawing.Size(176, 21);
			this.cmbFecha.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.cmbFecha.TabIndex = 1;
			this.cmbFecha.Value = new System.DateTime(2006, 10, 15, 0, 0, 0, 0);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(88, 24);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(36, 16);
			this.label1.TabIndex = 2;
			this.label1.Text = "Fecha";
			// 
			// btImpuestos
			// 
			this.btImpuestos.Location = new System.Drawing.Point(128, 248);
			this.btImpuestos.Name = "btImpuestos";
			this.btImpuestos.Size = new System.Drawing.Size(160, 23);
			this.btImpuestos.TabIndex = 3;
			this.btImpuestos.Text = "&Impuestos Nacionales";
			this.btImpuestos.Click += new System.EventHandler(this.btImpuestos_Click);
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
			// cmbCuentaCorr
			// 
			this.cmbCuentaCorr.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuentaCorr.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentaCorr.DataMember = "CuentaCorriente";
			this.cmbCuentaCorr.DataSource = this.cdsCuentaCorr;
			this.cmbCuentaCorr.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 6;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 9;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 4;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 5;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Width = 102;
			ultraGridColumn6.Header.VisiblePosition = 7;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 8;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 3;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 56;
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
																										 ultraGridColumn12});
			this.cmbCuentaCorr.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCuentaCorr.DisplayMember = "NumCuenta";
			this.cmbCuentaCorr.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCuentaCorr.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentaCorr.Location = new System.Drawing.Point(152, 56);
			this.cmbCuentaCorr.Name = "cmbCuentaCorr";
			this.cmbCuentaCorr.Size = new System.Drawing.Size(176, 21);
			this.cmbCuentaCorr.TabIndex = 4;
			this.cmbCuentaCorr.ValueMember = "idCuenta";
			// 
			// cdsCuentaCorr
			// 
			this.cdsCuentaCorr.BindingContextCtrl = this;
			this.cdsCuentaCorr.DataLibrary = "LibContabilidad";
			this.cdsCuentaCorr.DataLibraryUrl = "";
			this.cdsCuentaCorr.DataSetDef = "dsCuentaCorriente";
			this.cdsCuentaCorr.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCuentaCorr.Name = "cdsCuentaCorr";
			this.cdsCuentaCorr.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsCuentaCorr.SchemaDef = null;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(40, 56);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(90, 16);
			this.label2.TabIndex = 5;
			this.label2.Text = "Cuenta Corriente";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// btInternacional
			// 
			this.btInternacional.Location = new System.Drawing.Point(128, 280);
			this.btInternacional.Name = "btInternacional";
			this.btInternacional.Size = new System.Drawing.Size(160, 23);
			this.btInternacional.TabIndex = 6;
			this.btInternacional.Text = "&Internacional por Cobrar";
			this.btInternacional.Click += new System.EventHandler(this.btInternacional_Click);
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
			// txtNota
			// 
			this.txtNota.Location = new System.Drawing.Point(88, 96);
			this.txtNota.Multiline = true;
			this.txtNota.Name = "txtNota";
			this.txtNota.Size = new System.Drawing.Size(288, 104);
			this.txtNota.TabIndex = 7;
			this.txtNota.ValueChanged += new System.EventHandler(this.txtNota_ValueChanged);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(24, 96);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(28, 16);
			this.label3.TabIndex = 8;
			this.label3.Text = "Nota";
			// 
			// btCodigos
			// 
			this.btCodigos.Location = new System.Drawing.Point(360, 280);
			this.btCodigos.Name = "btCodigos";
			this.btCodigos.Size = new System.Drawing.Size(24, 24);
			this.btCodigos.TabIndex = 9;
			this.btCodigos.Text = "...";
			this.btCodigos.Click += new System.EventHandler(this.btCodigos_Click);
			// 
			// ultraButton1
			// 
			this.ultraButton1.Location = new System.Drawing.Point(128, 312);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.Size = new System.Drawing.Size(160, 23);
			this.ultraButton1.TabIndex = 10;
			this.ultraButton1.Text = "Impuestos Internacionales";
			this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// OmniLife
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(400, 358);
			this.Controls.Add(this.ultraButton1);
			this.Controls.Add(this.btCodigos);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.txtNota);
			this.Controls.Add(this.btInternacional);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbCuentaCorr);
			this.Controls.Add(this.btImpuestos);
			this.Controls.Add(this.cmbFecha);
			this.Controls.Add(this.btGanancia);
			this.Name = "OmniLife";
			this.Text = "OmniLife";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.OmniLife_Closing);
			this.Load += new System.EventHandler(this.OmniLife_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorr)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentaCorr)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNota)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		CultureInfo us = new CultureInfo("en-US");
		private void btGanancia_Click(object sender, System.EventArgs e)
		{
			#region Abre Archivo
			if (MessageBox.Show("¿Desea Subir Listado de Ganancias de Excel?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			string stArchivo = "";
			this.openFileDialog.Filter = "Excel (*.xls)|*.xls";
			this.openFileDialog.InitialDirectory = MenuLatinium.stDirInicio;
			if (DialogResult.OK == this.openFileDialog.ShowDialog())
			{
				stArchivo = this.openFileDialog.FileName.ToString();
			}
			else
			{
				MessageBox.Show("Cancelado por el usuario");
				Cursor = Cursors.Default;
				return;
			}

			stArchivo = stArchivo.ToUpper();
			#endregion Abre Archivo
			#region Lectura de titulos
			C1.C1Excel.C1XLBook miCompra = new C1.C1Excel.C1XLBook();
			try
			{
				miCompra.Load(@stArchivo);
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message);
				return;
			}
			XLSheet sheet = miCompra.Sheets[0];

			string [] stTitulo = new string[200];
			// Lectura de Titulos de columnas de Compras

			XLCell cell = sheet.GetCell(0, 0);
			StreamWriter sr= File.CreateText("LatiniumError.xls");
			int IdCliente = 0;
			string stSelect;
			string stAutFactura = "";
			string stAutImprenta = "";
			DateTime dtFechaCaduca = DateTime.Today;
			string stSerieFactura = "";
			#endregion Lectura de filas
			#region Lectura de Filas
			double dTotalAcumula = 0;
			for (int f = 0; f < sheet.Rows.Count; f++)
			{
				#region Tratamiento de Celda 0
				cell = sheet.GetCell(f, 0);
				if (cell == null || cell.Value == null) continue; // Linea Vacia
				#endregion Tratamiento de Celda 0		
				#region Verificacion de Cliente
				string stCodigoCliente = cell.Value.ToString().Trim();
				stSelect = "SELECT COUNT(*) FROM Cliente WHERE Codigo = '"
					+ stCodigoCliente + "'";

				int iCuenta = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
				if (iCuenta < 1) // No existe Cliente
				{
					cell = sheet.GetCell(f, 1);
					string stNombre = cell.Value.ToString().Trim();
					stNombre = stNombre.Replace("'", "''");
					stSelect = "Insert INTO Cliente (Codigo, Nombre, Proveedor, "
						+ " AutorizaFactura, AutorizaImprenta, FechaCaducidad, SerieFactura) "
						+ " VALUES ('" + stCodigoCliente + "', '" + stNombre+ "', 1, '"
						+ stAutFactura + "', '" + stAutImprenta + "', '"
						+ dtFechaCaduca.ToString("yyyyMMdd") + "', '"
						+ stSerieFactura + "')";
					Funcion.EjecutaSQL(cdsCompra, stSelect, true);
				}
				#endregion Verificacion de Cliente
				#region Identifiacion de Cliente
				stSelect = "Select idCliente FROM Cliente Where Codigo = '"
					+ stCodigoCliente + "'";
				IdCliente = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
				#endregion Identifiacionde Cliente
				#region Ingreso de Factura
				#region Lectura de Datos de Factura
				DateTime dtFecha = (DateTime) cmbFecha.Value;
				double dDescuento = 0;
				double dSubTotalIva0 = 0;
				double dSubTotalIva = 0;
				double dIva = 0;
				double dTotal = 0;
				try
				{
					cell = sheet.GetCell(f, 5);
					dSubTotalIva0 = (double) cell.Value;
					dTotal = dSubTotalIva0;
					dTotalAcumula += dTotal;
				}
				catch(System.Exception ex)
				{
					f++;
					MessageBox.Show(ex.Message, "Error en Fila: " + f.ToString());
					f--;
				}
				#endregion Lectura de Datos de Factura
				#region Select de Compra
				string stNumero = Funcion.NumeraFactura(cdsCompra, 14);
				stSelect = "Insert INTO Compra (idCliente, Fecha, ContadoCredito, "
					+ "SubTotalIva, SubTotalExcento, IVA, Total, Descuento, Numero, "
					+ "idTipoFactura, AutFactura, AutImprenta, SerieFactura, "
					+ "FechaCaducidad, idCredTributario, BienServicio, idComprobante)"
					+ "VALUES (" + IdCliente.ToString() + ", '" + dtFecha.ToString("yyyyMMdd")
					+ "', 2, " + dSubTotalIva.ToString("0.00", us) + ", " 
					+ dSubTotalIva0.ToString("0.00", us) + ", " 
					+ dIva.ToString("0.00", us) + ", " 
					+ dTotal.ToString("0.00", us) + ", " 
					+ dDescuento.ToString("0.00", us) + ", '" + stNumero + "', 14, '"
					+ stAutFactura + "', '" + stAutImprenta + "', '" + stSerieFactura 
					+ "', '" + dtFechaCaduca.ToString("yyyyMMdd") + "', 2, 2, 15)";
				Funcion.EjecutaSQL(cdsCompra, stSelect, true);
				#endregion Select de Compra
				#region IdCompra Generado
				stSelect = "Select TOP 1 IdCompra FROM Compra Order By IdCompra DESC";
				int IdCompra = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
				#endregion IdCompra Generado
				#region Ingreso de Detalle de Factura
				stSelect = "Insert INTO DetCompra (idArticulo, Cantidad, Precio, "
					+ "IdCompra,  Impuesto) "
					+ "VALUES (1, 1, " + dSubTotalIva0.ToString("0.00", us) 
					+ ", " + IdCompra.ToString() + ", 0)";
				Funcion.EjecutaSQL(cdsCompra, stSelect, true);
				#endregion Ingreso de Detalle de Factura
				#endregion Ingreso de Factura
			}
			#endregion Lectura de Filas
			#region Asiento de Ganancias
			DateTime dtFecha1 = (DateTime) cmbFecha.Value;
			string stNumero1 = Contabilidad.NumeroAsiento(cdsCompra, 1);
			stSelect = "Insert INTO Asiento (idTipoAsiento, Fecha, CodAsiento, "
				+ "Total, TotalH, Descripcion) "
				+ "VALUES (1, '" + dtFecha1.ToString("yyyyMMdd")
				+ "', '" + stNumero1 + "', "
				+ dTotalAcumula.ToString("0.00", us) + ", " 
				+ dTotalAcumula.ToString("0.00", us) + ", '" 
				+ "Ganancias')"; 
			Funcion.EjecutaSQL(cdsCompra, stSelect, true);

			stSelect = "Select TOP 1 idAsiento FROM Asiento Where CodAsiento = '"
				+ stNumero1 + "' Order By idAsiento DESC";
			int IdAsiento = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
			int IdCuenta = 0;

			// Ingreso de Detalle de Asiento
			IdCuenta = AsignaCuenta(14);
			DetalleAsiento(IdAsiento, IdCuenta, dTotalAcumula, 0);
			IdCuenta = AsignaCuenta(19);
			DetalleAsiento(IdAsiento, IdCuenta, 0, dTotalAcumula);

			#endregion Asiento de Ganancias
			MessageBox.Show("Archivo de Ganancias Subido", "Información");
		}

		// Ingreso de Total
		private void DetalleAsiento(int IdAsiento, int IdCuenta, double dDebe, double dHaber)
		{
			if (dDebe + dHaber < .00001) return;
			string stSelect = "Insert INTO AsientoDetalle (idAsiento, idCuenta, Debe, "
				+ "Haber) "
				+ "VALUES (" + IdAsiento.ToString() + ", " 
				+ IdCuenta.ToString() + ", " + dDebe.ToString("0.00", us)
				+ ", " + dHaber.ToString("0.00", us) + ")";
			Funcion.EjecutaSQL(cdsCompra, stSelect, true);
		}

		// Ingreso de Total
		private int AsignaCuenta(int IdAsigna)
		{
			string stSelect = "Select idCuenta FROM AsignaCuenta Where "
				+ " idAsignacion = " + IdAsigna.ToString();
			return Funcion.iEscalarSQL(cdsCompra, stSelect, true);
		}
		
		private void btImpuestos_Click(object sender, System.EventArgs e)
		{
			#region Abrir Archivo
			errorProvider.SetError(cmbCuentaCorr, "");
			if (cmbCuentaCorr.Text.Trim().Length == 0)
			{
				errorProvider.SetError(cmbCuentaCorr, "Ingrese Banco");
				return;
			}
			if (MessageBox.Show("¿Desea Subir Listado de Impuestos de Excel?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			string stArchivo = "";
			this.openFileDialog.Filter = "Excel (*.xls)|*.xls";
			this.openFileDialog.InitialDirectory = MenuLatinium.stDirInicio;
			if (DialogResult.OK == this.openFileDialog.ShowDialog())
			{
				stArchivo = this.openFileDialog.FileName.ToString();
			}
			else
			{
				MessageBox.Show("Cancelado por el usuario");
				this.Cursor = Cursors.Default;
				return;
			}
			this.Cursor = Cursors.WaitCursor;
			stArchivo = stArchivo.ToUpper();
			#endregion Abrir Archivo
			#region Lectura de Excel
			C1.C1Excel.C1XLBook miCompra = new C1.C1Excel.C1XLBook();
			try
			{
				miCompra.Load(@stArchivo);
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message);
				return;
			}
			XLSheet sheet = miCompra.Sheets[0];
			XLCell cell = sheet.GetCell(0, 0);
			StreamWriter sr= File.CreateText("LatiniumError.xls");
			#endregion Lectura de Excel
			#region Lectura de Filas
			int IdCliente = 0;
			string stSelect;
			double dTotalAcumula = 0;
			for (int f = 0; f < sheet.Rows.Count; f++)
			{
				#region Tratamiento de Celda 0
				cell = sheet.GetCell(f, 0);
				string stCodigoCliente = "";
				if (cell == null || cell.Value == null) continue;
				#endregion Tratamiento de Celda 0		
				#region Ingreso o creacion de Cliente
				stCodigoCliente = cell.Value.ToString().Trim();
				stSelect = "SELECT COUNT(*) FROM Cliente WHERE Codigo = '"
					+ stCodigoCliente + "'";
				int iCuenta = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
				if (iCuenta < 1) // No existe Cliente
				{
					cell = sheet.GetCell(f, 1);
					string stNombre = cell.Value.ToString().Trim();
					stSelect = "Insert INTO Cliente (Codigo, Nombre, Proveedor) "
						+ " VALUES ('" + stCodigoCliente + "', '" + stNombre+ "', 1)";
					Funcion.EjecutaSQL(cdsCompra, stSelect, true);
				}
				stSelect = "Select idCliente FROM Cliente Where Codigo = '"
					+ stCodigoCliente + "'";
				IdCliente = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
				#endregion Ingreso o creacion de Cliente		
				#region Lectura de Valores de Columnas dentro de fila f
				DateTime dtFecha = (DateTime) cmbFecha.Value;
				string stBeneficiario = "";
				string stCheque = "";
				string stTipo = "";
				double dComision = 0;
				double dIva = 0;
				double dRetRenta1 = 0;
				double dRetRenta8 = 0;
				double dRetIva70 = 0;
				double dRetIva100 = 0;
				double dRetRenta25 = 0;
				double dTotal = 0;
				string stSerieFactura = "";
				string stNumeroFact = "";
				string stAutFactura = "";
				string stAutImprenta = "";
				DateTime dtFechaCaduca = DateTime.Now;
				double dTotalH = 0;
				double dRetIva = 0;
				double dRetRenta = 0;
				try
				{
					stBeneficiario = sheet.GetCell(f, 1).Value.ToString();
					stCheque = sheet.GetCell(f, 2).Value.ToString();
					stTipo = sheet.GetCell(f, 3).Value.ToString();
					dComision = (double) sheet.GetCell(f, 4).Value;
					dTotalAcumula += dComision;
					dIva = (double) sheet.GetCell(f, 5).Value;
					dRetRenta1 = (double) sheet.GetCell(f, 6).Value;
					dRetRenta8 = (double) sheet.GetCell(f, 7).Value;
					dRetIva70 = (double) sheet.GetCell(f, 8).Value;
					dRetIva100 = (double) sheet.GetCell(f, 9).Value;
					dRetRenta25 = (double) sheet.GetCell(f, 10).Value;
					dTotal = (double) sheet.GetCell(f, 11).Value;
					stSerieFactura = sheet.GetCell(f, 12).Value.ToString();
					if (stSerieFactura.Length != 6)
					{
						int iLineaE = f + 1;
						MessageBox.Show("Error en serie de factura", "Error en linea " + iLineaE.ToString());
					}
					stNumeroFact = sheet.GetCell(f, 13).Value.ToString();
					stAutFactura = sheet.GetCell(f, 14).Value.ToString();
					stAutImprenta = sheet.GetCell(f, 15).Value.ToString();
					if (stAutFactura.Length != 10)
					{
						int iLineaE = f + 1;
						MessageBox.Show("Error en Autorizacion de factura", "Error en linea " + iLineaE.ToString());
					}
					if (sheet.GetCell(f, 16).GetType().ToString() == "STRING")
						dtFechaCaduca = (DateTime) sheet.GetCell(f, 16).Value;
					dTotalH = dComision + dIva;
					dRetIva = dRetIva100 + dRetIva70;
					dRetRenta = dRetRenta1 + dRetRenta8 + dRetRenta25;
				}
				catch(System.Exception ex)
				{
					f++;
					MessageBox.Show(ex.Message, "Error en: " + f.ToString());
					f--;
				}
				#endregion Lectura de columnas

				#region Creacion de Asiento

				string stNota = txtNota.Text.Trim();
				string stNumero = Contabilidad.NumeroAsiento(cdsCompra, 3);
				stSelect = "Insert INTO Asiento (idTipoAsiento, Fecha, CodAsiento, "
					+ "Total, TotalH, Descripcion, Cheque, Notas) "
					+ "VALUES (3, '" + dtFecha.ToString("yyyyMMdd")
					+ "', '" + stNumero + "', "
					+ dTotalH.ToString("0.00", us) + ", " 
					+ dTotalH.ToString("0.00", us) + ", '" 
					+ stBeneficiario + "', '" + stCheque + "', '"
					+ stNota + "')";
				Funcion.EjecutaSQL(cdsCompra, stSelect, true);

				stSelect = "Select TOP 1 idAsiento FROM Asiento Where CodAsiento = '"
					+ stNumero + "' Order By idAsiento DESC";
				int IdAsiento = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
				int IdCuenta = 0;

				// Ingreso de Detalle de Asiento
				IdCuenta = AsignaCuenta(19);
				DetalleAsiento(IdAsiento, IdCuenta, dComision, 0);

				IdCuenta = AsignaCuenta(11);
				DetalleAsiento(IdAsiento, IdCuenta, dIva, 0);

				string stPorcRI = "3"; // 70 % de Retencion de IVA
				if (stTipo.StartsWith("EXT")) stPorcRI = "4";
				stSelect = "Select idCuentaE FROM RetencionIva Where idRetIva = " + stPorcRI;
				IdCuenta = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
				DetalleAsiento(IdAsiento, IdCuenta, 0, dRetIva);

				string stPorcRR = "15"; // 1 % de Retencion de Renta
				if (stTipo.StartsWith("NAC")) stPorcRR = "1";
				if (stTipo.StartsWith("EXT")) stPorcRR = "36";
				if (dRetRenta1 > 0)
				{
					stSelect = "Select idCuentaE FROM ArticuloSRI Where idArticuloSRI = " + stPorcRR;
					IdCuenta = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
					DetalleAsiento(IdAsiento, IdCuenta, 0, dRetRenta1);
				}

				if (dRetRenta8 > 0)
				{
					stSelect = "Select idCuentaE FROM ArticuloSRI Where idArticuloSRI = " + stPorcRR;
					IdCuenta = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
					DetalleAsiento(IdAsiento, IdCuenta, 0, dRetRenta8);
				}

				if (dRetRenta25 > 0)
				{
					stSelect = "Select idCuentaE FROM ArticuloSRI Where idArticuloSRI = "+ stPorcRR;
					IdCuenta = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
					DetalleAsiento(IdAsiento, IdCuenta, 0, dRetRenta25);
				}

				IdCuenta = int.Parse(cmbCuentaCorr.Value.ToString());
				DetalleAsiento(IdAsiento, IdCuenta, 0, dTotal);
				#endregion Creacion de Asiento

				#region Creacion de Factura de Compra
				stSelect = "Insert INTO Compra (idCliente, Fecha, ContadoCredito, "
					+ "SubTotalIva, SubTotalExcento, IVA, Total, Descuento, Numero, "
					+ "idTipoFactura, AutFactura, AutImprenta, SerieFactura, "
					+ "FechaCaducidad, idCredTributario, BienServicio, idComprobante, idAsiento)"
					+ "VALUES (" + IdCliente.ToString() + ", '" + dtFecha.ToString("yyyyMMdd")
					+ "', 1, " + dComision.ToString("0.00", us) + ", 0, " 
					+ dIva.ToString("0.00", us) + ", " 
					+ dTotalH.ToString("0.00", us) + ", 0" 
					+ ", '" + stNumeroFact + "', 4, '"
					+ stAutFactura + "', '" + stAutImprenta + "', '" + stSerieFactura 
					+ "', '" + dtFechaCaduca.ToString("yyyyMMdd") + "', 1, 2, 1, "
					+ IdAsiento.ToString() +")";
				Funcion.EjecutaSQL(cdsCompra, stSelect, true);

				stSelect = "Select TOP 1 IdCompra FROM Compra Order By IdCompra DESC";
				int IdCompra = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
				// Ingreso de Detalle de Factura
				int IdTipoAsi = 1; //Juridico
//				if (stTipo.StartsWith("NAC")) IdTipoAsi = 2;
//				if (stTipo.StartsWith("EXT")) IdTipoAsi = 3;

				stSelect = "Insert INTO DetCompra (idArticulo, Cantidad, Precio, "
					+ "IdCompra,  Impuesto) "
					+ "VALUES (" + IdTipoAsi.ToString() + ", 1, " + dComision.ToString("0.00", us) 
					+ ", " + IdCompra.ToString() + ", 12)";
				Funcion.EjecutaSQL(cdsCompra, stSelect, true);
				#endregion Creacion de factura de compra
				#region Creacion de Retencion
				// Retenciones debe variar para adaptarse a las nuevas retenciones
				string stNumeroRet = Contabilidad.NumeroAsiento(cdsCompra, 7);
				string stSerieRet = drSeteo["SriSerieRet"].ToString();
				string stAutorizaRet = drSeteo["SriAutRet"].ToString();
				dRetRenta = dRetRenta1 + dRetRenta8 + dRetRenta25;

				stSelect = "Select idCuentaCorriente FROM CuentaCorriente "
					+ "Where idCuenta = " + IdCuenta.ToString();
				IdCuenta = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
				stSelect = "Insert Into Retencion (idAsiento, idCliente, idTipoRetencion, "
					+ "PorcIva, FacturaIva121, IvaTotal1, TotalFactura, Total, "
					+ "CodAsiento, idRetRenta, idRetIva1, RentaRetenida, IvaRetenido1, "
					+ "CodigoSri, Fecha, Codigo, Beneficiario, idCuentaCorriente, "
					+ "Cheque, IngresoEgreso, IdCompra, AutorizaRet, SerieRet, NumFactura, Notas)"
					+ "Values (" + IdAsiento.ToString() + ", " + IdCliente.ToString() 
					+ ", 1, 12, " + dComision.ToString("0.00", us) 
					+ ", " + dIva.ToString("0.00", us) + ", " 
					+ dTotalH.ToString("0.00", us)
					+ ", " + dTotal.ToString("0.00", us) + ", '" + stNumero + "', "
					+ stPorcRR + ", " + stPorcRI + ", " + dRetRenta.ToString("0.00", us)
					+ ", " + dRetIva.ToString("0.00", us) + ", 303, '" 
					+ dtFecha.ToString("yyyyMMdd") + "', '" + stNumeroRet + "', '" 
					+ stBeneficiario + "', " + IdCuenta.ToString()
					+ ", " + stCheque + ",2, " + IdCompra.ToString() + ", '"
					+ stAutorizaRet + "', '" + stSerieRet + "', '" + stSerieFactura + stNumeroFact
					+ "', '" + txtNota.Text.Trim() + "')";
				Funcion.EjecutaSQL(cdsCompra, stSelect, true);

				int IdRetArtSri = 0;
				if (dRetRenta1 > 0) IdRetArtSri = 15;
				if (dRetRenta8 > 0) IdRetArtSri = 1;
				if (dRetRenta25 > 0) IdRetArtSri = 3; //  Cambiar a 36 si se necesita el codigo 411
				stSelect = "Select Top 1 idRetencion From Retencion Order by idRetencion DESC";
				int IdRet = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
				stSelect = "Insert into RetDetalle (idRetencion, idRetRenta, Base, Retenido) "
					+ "Values (" + IdRet.ToString() + ", " + IdRetArtSri.ToString() 
					+ ", " + dComision.ToString("0.00", us)
					+ ", " + dRetRenta.ToString("0.00", us) + ")";
				Funcion.EjecutaSQL(cdsCompra, stSelect, true);
				#endregion creacion de retencion
				#region Generacion de Pago va cancelando todas las facturas que esten pendientes
				stSelect = "Exec OmniLifePago '" + stCodigoCliente + "', "
					+ dComision.ToString("0.00", us) + ", '" + stNumero  + "', '"
					+ dtFecha.ToString("yyyyMMdd") + "'";
				Funcion.EjecutaSQL(cdsCompra, stSelect, true);
				#endregion Generacion de Pago Va cancelando todas las facturas que esten pendientes
			}
			#endregion Lectura de Filas
			#region Asiento de resumen
			DateTime dtFecha1 = (DateTime) cmbFecha.Value;
			string stNumero1 = Contabilidad.NumeroAsiento(cdsCompra, 1);
			stSelect = "Insert INTO Asiento (idTipoAsiento, Fecha, CodAsiento, "
				+ "Total, TotalH, Descripcion) "
				+ "VALUES (1, '" + dtFecha1.ToString("yyyyMMdd")
				+ "', '" + stNumero1 + "', "
				+ dTotalAcumula.ToString("0.00", us) + ", " 
				+ dTotalAcumula.ToString("0.00", us) + ", '" 
				+ "Ganancias')"; 
			Funcion.EjecutaSQL(cdsCompra, stSelect, true);

			stSelect = "Select TOP 1 idAsiento FROM Asiento Where CodAsiento = '"
				+ stNumero1 + "' Order By idAsiento DESC";
			int IdAsiento1 = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
			int IdCuenta1 = 0;

			// Ingreso de Detalle de Asiento
			IdCuenta1 = AsignaCuenta(12);
			DetalleAsiento(IdAsiento1, IdCuenta1, dTotalAcumula, 0);
			IdCuenta1 = AsignaCuenta(14);
			DetalleAsiento(IdAsiento1, IdCuenta1, 0, dTotalAcumula);
			#endregion Asiento de resumen
			MessageBox.Show("Archivo de Impuestos Subido", "Información");		
		}

		private void cdsCompra_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCompra.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeteoF.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		C1.Data.C1DataRow drSeteo;
		private void OmniLife_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'OmniLife'";
			Funcion.EjecutaSQL(cdsCompra, stAudita, true);
			cmbFecha.Value = DateTime.Today;
			drSeteo = cdsSeteo.TableViews["Seteo"].Rows[0];
		}

		private void btInternacional_Click(object sender, System.EventArgs e)
		{
			// Debe ser como el segundo proceso
			#region Abrir Archivo
			errorProvider.SetError(cmbCuentaCorr, "");
			if (cmbCuentaCorr.Text.Trim().Length == 0)
			{
				errorProvider.SetError(cmbCuentaCorr, "Ingrese Banco");
				return;
			}
			if (MessageBox.Show("¿Desea Subir Listado de Cheques Internacionales de Excel?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			string stArchivo = "";
			this.openFileDialog.Filter = "Excel (*.xls)|*.xls";
			this.openFileDialog.InitialDirectory = MenuLatinium.stDirInicio;
			if (DialogResult.OK == this.openFileDialog.ShowDialog())
			{
				stArchivo = this.openFileDialog.FileName.ToString();
			}
			else
			{
				MessageBox.Show("Cancelado por el usuario");
				this.Cursor = Cursors.Default;
				return;
			}
			this.Cursor = Cursors.WaitCursor;
			stArchivo = stArchivo.ToUpper();
			#endregion Abrir Archivo
			#region Lectura de Excel
			C1.C1Excel.C1XLBook miCompra = new C1.C1Excel.C1XLBook();
			try
			{
				miCompra.Load(@stArchivo);
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message);
				return;
			}
			XLSheet sheet = miCompra.Sheets[0];
			XLCell cell = sheet.GetCell(0, 0);
			StreamWriter sr= File.CreateText("LatiniumError.xls");
			#endregion Lectura de Excel
			#region Lectura de Filas
			int IdCliente = 0;
			string stSelect;
			double dTotalAcumula = 0;
			for (int f = 0; f < sheet.Rows.Count; f++)
			{
				#region Tratamiento de Celda 0
				cell = sheet.GetCell(f, 0);
				string stCodigoCliente = "";
				if (cell == null || cell.Value == null) continue;
				#endregion Tratamiento de Celda 0		
				#region Ingreso o creacion de Cliente
				stCodigoCliente = cell.Value.ToString().Trim();
				stSelect = "SELECT COUNT(*) FROM Cliente WHERE Codigo = '"
					+ stCodigoCliente + "'";
				int iCuenta = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
				if (iCuenta < 1) // No existe Cliente
				{
					cell = sheet.GetCell(f, 1);
					string stNombre = cell.Value.ToString().Trim();
					stSelect = "Insert INTO Cliente (Codigo, Nombre, Proveedor) "
						+ " VALUES ('" + stCodigoCliente + "', '" + stNombre+ "', 1)";
					Funcion.EjecutaSQL(cdsCompra, stSelect, true);
				}
				stSelect = "Select idCliente FROM Cliente Where Codigo = '"
					+ stCodigoCliente + "'";
				IdCliente = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
				#endregion Ingreso o creacion de Cliente		
				#region Lectura de Valores de Columnas dentro de fila f
				DateTime dtFecha = (DateTime) cmbFecha.Value;
				string stBeneficiario = "";
				string stCheque = "";
				string stTipo = "";
				double dComision = 0;
				double dTotal = 0;
				try
				{
					stBeneficiario = sheet.GetCell(f, 1).Value.ToString();
					stCheque = sheet.GetCell(f, 9).Value.ToString();
					stTipo = sheet.GetCell(f, 2).Value.ToString();
					dComision = (double) sheet.GetCell(f, 8).Value;
					dTotalAcumula += dComision;
					dTotal = (double) sheet.GetCell(f, 8).Value;
				}
				catch(System.Exception ex)
				{
					f++;
					MessageBox.Show(ex.Message, "Error en: " + f.ToString());
					f--;
				}
				#endregion Lectura de columnas
				#region Creacion de Factura de Compra
//				Comentarios generados de toda seccion porque solo se genera la parte contabla
//				stSelect = "Insert INTO Compra (idCliente, Fecha, ContadoCredito, "
//					+ "SubTotalIva, SubTotalExcento, IVA, Total, Descuento, Numero, "
//					+ "idTipoFactura, "
//					+ "idCredTributario, BienServicio, idComprobante)"
//					+ "VALUES (" + IdCliente.ToString() + ", '" + dtFecha.ToString("yyyyMMdd")
//					+ "', 1, " + dComision.ToString("0.00", us) + ", 0, " 
//					+ "0, " 
//					+ dComision.ToString("0.00", us) + ", 0" 
//					+ ", '" + stCheque + "', 4"
//					+ ", 1, 2, 1)";
//				Funcion.EjecutaSQL(cdsCompra, stSelect, true);
//
//				stSelect = "Select TOP 1 IdCompra FROM Compra Order By IdCompra DESC";
//				int IdCompra = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
//				// Ingreso de Detalle de Factura
//
//				stSelect = "Insert INTO DetCompra (idArticulo, Cantidad, Precio, "
//					+ "IdCompra,  Impuesto) "
//					+ "VALUES (1, 1, " + dComision.ToString("0.00", us) 
//					+ ", " + IdCompra.ToString() + ", 12)";
//				Funcion.EjecutaSQL(cdsCompra, stSelect, true);
				#endregion Creacion de factura de compra
				#region Creacion de Asiento

				string stNota = txtNota.Text.Trim();
				string stNumero = Contabilidad.NumeroAsiento(cdsCompra, 3);
				stSelect = "Insert INTO Asiento (idTipoAsiento, Fecha, CodAsiento, "
					+ "Total, TotalH, Descripcion, Cheque, Notas, Numero) "
					+ "VALUES (3, '" + dtFecha.ToString("yyyyMMdd")
					+ "', '" + stNumero + "', "
					+ dTotal.ToString("0.00", us) + ", " 
					+ dTotal.ToString("0.00", us) + ", '" 
					+ stBeneficiario + "', '" + stCheque + "', '"
					+ stNota + "', '" +  stCodigoCliente + "')";
				Funcion.EjecutaSQL(cdsCompra, stSelect, true);

				stSelect = "Select TOP 1 idAsiento FROM Asiento Where CodAsiento = '"
					+ stNumero + "' Order By idAsiento DESC";
				int IdAsiento = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
				int IdCuenta = 0;

				// Ingreso de Detalle de Asiento
				string stCodInt = sheet.GetCell(f, 2).Value.ToString();
				stSelect = "Select IsNull(idCuenta2, 0) From OmniLifeCodigo where Codigo = '"
					+ stCodInt + "'";
				IdCuenta = Funcion.iEscalarSQL(cdsCompra, stSelect, false);
				DetalleAsiento(IdAsiento, IdCuenta, dComision, 0);

				IdCuenta = int.Parse(cmbCuentaCorr.Value.ToString());
				DetalleAsiento(IdAsiento, IdCuenta, 0, dTotal);
				#endregion Creacion de Asiento
				#region Generacion de Pago va cancelando todas las facturas que esten pendientes
//				stSelect = "Exec OmniLifePago '" + stCodigoCliente + "', "
//					+ dComision.ToString("0.00", us) + ", '" + stNumero  + "', '"
//					+ dtFecha.ToString("yyyyMMdd") + "'";
//				Funcion.EjecutaSQL(cdsCompra, stSelect, true);
				#endregion Generacion de Pago Va cancelando todas las facturas que esten pendientes
			}
			#endregion Lectura de Filas
			#region Asiento de resumen anterior
			/*
			DateTime dtFecha1 = (DateTime) cmbFecha.Value;
			string stNumero1 = Contabilidad.NumeroAsiento(cdsCompra, 1);
			stSelect = "Insert INTO Asiento (idTipoAsiento, Fecha, CodAsiento, "
				+ "Total, TotalH, Descripcion) "
				+ "VALUES (1, '" + dtFecha1.ToString("yyyyMMdd")
				+ "', '" + stNumero1 + "', "
				+ dTotalAcumula.ToString("0.00", us) + ", " 
				+ dTotalAcumula.ToString("0.00", us) + ", '" 
				+ "Impuesto internacional')"; 
			Funcion.EjecutaSQL(cdsCompra, stSelect, true);

			stSelect = "Select TOP 1 idAsiento FROM Asiento Where CodAsiento = '"
				+ stNumero1 + "' Order By idAsiento DESC";
			int IdAsiento1 = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
			int IdCuenta1 = 0;

			// Ingreso de Detalle de Asiento
			IdCuenta1 = AsignaCuenta(12);
			DetalleAsiento(IdAsiento1, IdCuenta1, dTotalAcumula, 0);
			IdCuenta1 = AsignaCuenta(14);
			DetalleAsiento(IdAsiento1, IdCuenta1, 0, dTotalAcumula);
			*/
			#endregion Asiento de resumen
			this.Cursor = Cursors.Default;
			MessageBox.Show("Archivo de Internacionales por cobrar Subido", "Información");
		}

		private void timer1_Tick(object sender, System.EventArgs e)
		{
		}

		private void btCodigos_Click(object sender, System.EventArgs e)
		{
			OmniLifeCargo miCodigo = new OmniLifeCargo();
			miCodigo.MdiParent = this.MdiParent;
			miCodigo.Show();
		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			errorProvider.SetError(cmbCuentaCorr, "");
			if (cmbCuentaCorr.Text.Trim().Length == 0)
			{
				errorProvider.SetError(cmbCuentaCorr, "Ingrese Banco");
				return;
			}
			#region Abre Archivo
			if (MessageBox.Show("¿Desea Subir Listado de Ganancias Internacionales de Excel?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			string stArchivo = "";
			this.openFileDialog.Filter = "Excel (*.xls)|*.xls";
			this.openFileDialog.InitialDirectory = MenuLatinium.stDirInicio;
			if (DialogResult.OK == this.openFileDialog.ShowDialog())
			{
				stArchivo = this.openFileDialog.FileName.ToString();
			}
			else
			{
				MessageBox.Show("Cancelado por el usuario");
				Cursor = Cursors.Default;
				return;
			}

			stArchivo = stArchivo.ToUpper();
			#endregion Abre Archivo
			#region Lectura de titulos
			C1.C1Excel.C1XLBook miCompra = new C1.C1Excel.C1XLBook();
			try
			{
				miCompra.Load(@stArchivo);
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message);
				return;
			}
			XLSheet sheet = miCompra.Sheets[0];

			string [] stTitulo = new string[200];
			// Lectura de Titulos de columnas de Compras

			XLCell cell = sheet.GetCell(0, 0);
			StreamWriter sr= File.CreateText("LatiniumError.xls");
			int IdCliente = 0;
			string stSelect;
			string stAutFactura = "";
			string stAutImprenta = "";
			DateTime dtFechaCaduca = DateTime.Today;
			string stSerieFactura = "";
			#endregion Lectura de filas
			#region Lectura de Filas
			double dTotalAcumula = 0;
			DateTime dtFecha1 = (DateTime) cmbFecha.Value;
			for (int f = 0; f < sheet.Rows.Count; f++)
			{
				#region Tratamiento de Celda 0
				cell = sheet.GetCell(f, 0);
				if (cell == null || cell.Value == null) continue; // Linea Vacia
				#endregion Tratamiento de Celda 0		
				#region Verificacion de Cliente
				string stCodigoCliente = cell.Value.ToString().Trim();
				stSelect = "SELECT COUNT(*) FROM Cliente WHERE Codigo = '"
					+ stCodigoCliente + "'";

				int iCuenta = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
				if (iCuenta < 1) // No existe Cliente
				{
					cell = sheet.GetCell(f, 4);
					string stNombre = cell.Value.ToString().Trim();
					stNombre = stNombre.Replace("'", "''");
					stSelect = "Insert INTO Cliente (Codigo, Nombre, Proveedor) "
						+ " VALUES ('" + stCodigoCliente + "', '" + stNombre+ "', 1)";
					Funcion.EjecutaSQL(cdsCompra, stSelect, true);
				}
				#endregion Verificacion de Cliente
				#region Identificacion de Cliente
				stSelect = "Select idCliente FROM Cliente Where Codigo = '"
					+ stCodigoCliente + "'";
				IdCliente = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
				#endregion Identificacion de Cliente
				#region Ingreso de Asiento y Factura
				#region Lectura de Datos de Factura
				DateTime dtFecha = (DateTime) cmbFecha.Value;
				double dDescuento = 0;
				double dSubTotalIva0 = 0;
				double dSubTotalIva = 0;
				double dIva = 0;
				double dTotal = 0;
				try
				{
					cell = sheet.GetCell(f, 5);
					dSubTotalIva0 = (double) cell.Value;
					dTotal = dSubTotalIva0;
					dTotalAcumula += dTotal;
				}
				catch(System.Exception ex)
				{
					f++;
					MessageBox.Show(ex.Message, "Error en Fila: " + f.ToString());
					f--;
				}
				#endregion Lectura de Datos de Factura
				#region Asiento de Linea
				string stNumeroL = Contabilidad.NumeroAsiento(cdsCompra, 1);

				cell = sheet.GetCell(f, 4);
				string stNombreCli = cell.Value.ToString().Trim();
				string stNotas = txtNota.Text.Trim();

				stSelect = "Insert INTO Asiento (idTipoAsiento, Fecha, CodAsiento, "
					+ "Total, TotalH, Descripcion, Numero, Notas) "
					+ "VALUES (1, '" + dtFecha1.ToString("yyyyMMdd")
					+ "', '" + stNumeroL + "', "
					+ dSubTotalIva0.ToString("0.00", us) + ", " 
					+ dSubTotalIva0.ToString("0.00", us) + ", '" 
					+ stNombreCli + "', '" + stCodigoCliente
					+ "', '" + stNotas + "')"; 
				Funcion.EjecutaSQL(cdsCompra, stSelect, true);

				stSelect = "Select TOP 1 idAsiento FROM Asiento Where CodAsiento = '"
					+ stNumeroL + "' Order By idAsiento DESC";
				int IdAsientoL = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
				int IdCuentaL = 0;
				// Ingreso de Detalle de Asiento
				// Cuenta del 25%
				stSelect = "Select IsNull(idCuenta, 0) From Cuenta where Codigo = '4.3.1.1.1.90'";
				IdCuentaL = Funcion.iEscalarSQL(cdsCompra, stSelect, false);
				DetalleAsiento(IdAsientoL, IdCuentaL, dSubTotalIva0, 0);

				//Codigo Internacional
				string stCodInt = sheet.GetCell(f, 1).Value.ToString();
				stSelect = "Select IsNull(idCuenta1, 0) From OmniLifeCodigo where Codigo = '"
					+ stCodInt + "'";
				IdCuentaL = Funcion.iEscalarSQL(cdsCompra, stSelect, false);
				double dCheque = (double) sheet.GetCell(f, 8).Value;
				DetalleAsiento(IdAsientoL, IdCuentaL, 0, dCheque);

				//Impuesto
				stSelect = "Select IsNull(idCuentaE, 0) From ArticuloSRI where Codigo = '421'";
				IdCuentaL = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
				double dImpuesto = (double) sheet.GetCell(f, 6).Value;
				DetalleAsiento(IdAsientoL, IdCuentaL, 0, dImpuesto);
				#endregion Asiento de Linea
				#region Select de Compra
				string stNumero = Funcion.NumeraFactura(cdsCompra, 14);
				stSelect = "Insert INTO Compra (idCliente, Fecha, ContadoCredito, "
					+ "SubTotalIva, SubTotalExcento, IVA, Total, Descuento, Numero, "
					+ "idTipoFactura, AutFactura, AutImprenta, SerieFactura, "
					+ "FechaCaducidad, idCredTributario, BienServicio, idComprobante, "
					+ "idAsiento) "
					+ "VALUES (" + IdCliente.ToString() + ", '" + dtFecha.ToString("yyyyMMdd")
					+ "', 1, " + dSubTotalIva.ToString("0.00", us) + ", " 
					+ dSubTotalIva0.ToString("0.00", us) + ", " 
					+ dIva.ToString("0.00", us) + ", " 
					+ dTotal.ToString("0.00", us) + ", " 
					+ dDescuento.ToString("0.00", us) + ", '" + stNumero + "', 4, '"
					+ stAutFactura + "', '" + stAutImprenta + "', '" + stSerieFactura 
					+ "', '" + dtFechaCaduca.ToString("yyyyMMdd") + "', 2, 2, 12, " + 
					IdAsientoL.ToString() + ")";
				//				antes generaba idtipofactura = 14 y comprobante 20 no aplica cambia a 12 doc ext
				Funcion.EjecutaSQL(cdsCompra, stSelect, true);
				#endregion Select de Compra
				#region IdCompra Generado
				stSelect = "Select TOP 1 IdCompra FROM Compra Order By IdCompra DESC";
				int IdCompra = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
				#endregion IdCompra Generado
				#region Ingreso de Detalle de Factura
				stSelect = "Insert INTO DetCompra (idArticulo, Cantidad, Precio, "
					+ "IdCompra,  Impuesto) "
					+ "VALUES (1, 1, " + dSubTotalIva0.ToString("0.00", us) 
					+ ", " + IdCompra.ToString() + ", 0)";
				Funcion.EjecutaSQL(cdsCompra, stSelect, true);
				#endregion Ingreso de Detalle de Factura

				#region Creacion de Retencion
				// Retenciones debe variar para adaptarse a las nuevas retenciones
				string stNumeroRet = Contabilidad.NumeroAsiento(cdsCompra, 7);
				string stSerieRet = drSeteo["SriSerieRet"].ToString();
				string stAutorizaRet = drSeteo["SriAutRet"].ToString();

				double dComision = dSubTotalIva0;
				double dRetRenta = dSubTotalIva0 / 4.0;
				int IdCuenta1 = (int) cmbCuentaCorr.Value;
				stSelect = "Select idCuentaCorriente FROM CuentaCorriente "
					+ "Where idCuenta = " + IdCuenta1.ToString();
				IdCuenta1 = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
				string stPorcRR = "36";
				string stBeneficiario = sheet.GetCell(f, 4).Value.ToString();
				string stCheque = "";
				string stNumeroFact = "";
				stSelect = "Insert Into Retencion (idAsiento, idCliente, idTipoRetencion, "
					+ "PorcIva, FacturaIva121, IvaTotal1, TotalFactura, Total, "
					+ "CodAsiento, idRetRenta, idRetIva1, RentaRetenida, IvaRetenido1, "
					+ "CodigoSri, Fecha, Codigo, Beneficiario, idCuentaCorriente, "
					+ "Cheque, IngresoEgreso, IdCompra, AutorizaRet, SerieRet, NumFactura, Notas)"
					+ "Values (" + IdAsientoL.ToString() + ", " + IdCliente.ToString() 
					+ ", 1, 0, " + dComision.ToString("0.00", us) 
					+ ", " + dIva.ToString("0.00", us) + ", " 
					+ dSubTotalIva0.ToString("0.00", us)
					+ ", " + dTotal.ToString("0.00", us) + ", '" + stNumero + "', "
					+ stPorcRR + ", 0, " + dRetRenta.ToString("0.00", us)
					+ ", 0, 303, '" 
					+ dtFecha.ToString("yyyyMMdd") + "', '" + stNumeroRet + "', '" 
					+ stBeneficiario + "', " + IdCuenta1.ToString()
					+ ", '" + stCheque + "',2, " + IdCompra.ToString() + ", '"
					+ stAutorizaRet + "', '" + stSerieRet + "', '" + stNumeroFact
					+ "', '" + txtNota.Text.Trim() + "')";
				Funcion.EjecutaSQL(cdsCompra, stSelect, true);

				int IdRetArtSri = 36;
				stSelect = "Select Top 1 idRetencion From Retencion Order by idRetencion DESC";
				int IdRet = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
				stSelect = "Insert into RetDetalle (idRetencion, idRetRenta, Base, Retenido) "
					+ "Values (" + IdRet.ToString() + ", " + IdRetArtSri.ToString() 
					+ ", " + dComision.ToString("0.00", us)
					+ ", " + dRetRenta.ToString("0.00", us) + ")";
				Funcion.EjecutaSQL(cdsCompra, stSelect, true);
				#endregion creacion de retencion
				#region Generacion de Pago va cancelando todas las facturas que esten pendientes
				stSelect = "Exec OmniLifePago '" + stCodigoCliente + "', "
					+ dComision.ToString("0.00", us) + ", '" + stNumero  + "', '"
					+ dtFecha.ToString("yyyyMMdd") + "'";
				Funcion.EjecutaSQL(cdsCompra, stSelect, true);
				#endregion Generacion de Pago Va cancelando todas las facturas que esten pendientes

				#endregion Ingreso de Factura
			}
			#endregion Lectura de Filas
			#region Asiento de Ganancias
			string stNumero1 = Contabilidad.NumeroAsiento(cdsCompra, 1);
			stSelect = "Insert INTO Asiento (idTipoAsiento, Fecha, CodAsiento, "
				+ "Total, TotalH, Descripcion) "
				+ "VALUES (1, '" + dtFecha1.ToString("yyyyMMdd")
				+ "', '" + stNumero1 + "', "
				+ dTotalAcumula.ToString("0.00", us) + ", " 
				+ dTotalAcumula.ToString("0.00", us) + ", '" 
				+ "Ganancias Internacionales')"; 
			Funcion.EjecutaSQL(cdsCompra, stSelect, true);

			stSelect = "Select TOP 1 idAsiento FROM Asiento Where CodAsiento = '"
				+ stNumero1 + "' Order By idAsiento DESC";
			int IdAsiento = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
			int IdCuenta = 0;

			// Ingreso de Detalle de Asiento
			IdCuenta = AsignaCuenta(19);
			DetalleAsiento(IdAsiento, IdCuenta, dTotalAcumula, 0);
			IdCuenta = AsignaCuenta(14);
			DetalleAsiento(IdAsiento, IdCuenta, 0, dTotalAcumula);

			#endregion Asiento de Ganancias
			MessageBox.Show("Archivo de Impuestos Internacional Subido", "Información");
		}

		private void txtNota_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void OmniLife_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'OmniLife'";
			Funcion.EjecutaSQL(cdsCompra, stAudita, true);
		}
	}
}
