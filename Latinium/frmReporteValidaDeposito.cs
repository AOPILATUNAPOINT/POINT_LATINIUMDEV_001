using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteValidaDeposito.
	/// </summary>
	public class frmReporteValidaDeposito : System.Windows.Forms.Form
	{
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReporteValidaDeposito()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		public Infragistics.Win.UltraWinGrid.UltraGrid uGridDepositos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource8;
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Button btnCancelar;
		private C1.Data.C1DataSet cdsSeteoF;
    decimal dValor = 0;
		DateTime dtFecha;
		private System.Windows.Forms.Label lblEstado;
		int iBodega = 0;
		int idDetalleCaja = 0;
		decimal dEfectivoCaja = 0;
		int idCaja = 0;
		public frmReporteValidaDeposito(decimal DValor, DateTime DtFecha, int IBodega, int IdDetalleCaja, decimal DEfectivoCaja, int IdCaja )
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			dValor = DValor;
			iBodega = IBodega;
			dtFecha = DtFecha;
			idDetalleCaja = IdDetalleCaja;
      dEfectivoCaja = DEfectivoCaja;
			idCaja = IdCaja;
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("R");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Origen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Factura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RucCedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Efectivo");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Utilizado");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Disponible");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 10, true);
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Efectivo", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Efectivo", 6, true);
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Disponible", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Disponible", 9, true);
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Utilizado", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Utilizado", 8, true);
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("R");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Origen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Factura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("RucCedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Efectivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Utilizado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Disponible");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			this.uGridDepositos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource8 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.lblEstado = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.uGridDepositos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridDepositos
			// 
			this.uGridDepositos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDepositos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDepositos.DataSource = this.ultraDataSource8;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridDepositos.DisplayLayout.Appearance = appearance1;
			this.uGridDepositos.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.Caption = "...";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 30;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 78;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 80;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.Caption = "Número";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 80;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "R.U.C.-Cédula";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 100;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 257;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance2;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 80;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 76;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance3;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 80;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance4;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 80;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance5;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 80;
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
																										 ultraGridColumn11});
			ultraGridBand1.Header.Caption = "Depositos";
			ultraGridBand1.HeaderVisible = true;
			ultraGridBand1.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			ultraGridBand1.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance6;
			summarySettings1.DisplayFormat = "{0:#,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance7;
			summarySettings2.DisplayFormat = "{0:#,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance8;
			summarySettings3.DisplayFormat = "{0:#,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance9;
			summarySettings4.DisplayFormat = "{0:#,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4});
			ultraGridBand1.SummaryFooterCaption = "Total Depositos";
			this.uGridDepositos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridDepositos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			appearance10.BackColor = System.Drawing.Color.Transparent;
			this.uGridDepositos.DisplayLayout.Override.CardAreaAppearance = appearance10;
			appearance11.ForeColor = System.Drawing.Color.Black;
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDepositos.DisplayLayout.Override.CellAppearance = appearance11;
			this.uGridDepositos.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance12.FontData.BoldAsString = "True";
			appearance12.FontData.Name = "Arial";
			appearance12.FontData.SizeInPoints = 8F;
			appearance12.ForeColor = System.Drawing.Color.White;
			appearance12.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDepositos.DisplayLayout.Override.HeaderAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDepositos.DisplayLayout.Override.RowSelectorAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDepositos.DisplayLayout.Override.SelectedRowAppearance = appearance14;
			this.uGridDepositos.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDepositos.Location = new System.Drawing.Point(8, 48);
			this.uGridDepositos.Name = "uGridDepositos";
			this.uGridDepositos.Size = new System.Drawing.Size(888, 272);
			this.uGridDepositos.TabIndex = 251;
			this.uGridDepositos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridDepositos_KeyDown);
			this.uGridDepositos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDepositos_AfterCellUpdate);
			this.uGridDepositos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridDepositos_InitializeLayout);
			// 
			// ultraDataSource8
			// 
			ultraDataColumn1.DataType = typeof(bool);
			ultraDataColumn7.DataType = typeof(System.Decimal);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn11.DataType = typeof(System.Decimal);
			this.ultraDataSource8.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn11});
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.Location = new System.Drawing.Point(16, 328);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(88, 23);
			this.btnAceptar.TabIndex = 252;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Location = new System.Drawing.Point(112, 328);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(88, 23);
			this.btnCancelar.TabIndex = 253;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.Visible = false;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
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
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.lblEstado.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(16, 8);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 23);
			this.lblEstado.TabIndex = 254;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// frmReporteValidaDeposito
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(904, 358);
			this.ControlBox = false;
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.uGridDepositos);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.Name = "frmReporteValidaDeposito";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Movimientos De Caja En Efectivo";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReporteValidaDeposito_KeyDown);
			this.Load += new System.EventHandler(this.frmReporteValidaDeposito_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridDepositos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		decimal dDepositadoHastaahora = 0;
		private void frmReporteValidaDeposito_Load(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Exec ReporteValidaDeposito '{0}', {1}", Convert.ToDateTime(dtFecha).ToString("yyyyMMdd"), iBodega);
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridDepositos);

			dDepositadoHastaahora = Funcion.decEscalarSQL(cdsSeteoF, string.Format(" select isnull (sum (Valor), 0) from dbo.ValidaDeposito where idCaja = {0} ",
				idCaja));

			int idCuponesWeb = 0;
			idCuponesWeb = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" select count (1) from dbo.ValidaDeposito where idDetalleCaja = {0} ",
				idDetalleCaja));

			decimal dValordeposito = 0;
      dValordeposito = dValor;
			if ( idCuponesWeb > 0 )
			{
				decimal dValoresV = 0;
				dValoresV = Funcion.decEscalarSQL(cdsSeteoF, string.Format(" select isnull(sum (Valor),0) from dbo.ValidaDeposito where idDetalleCaja = {0} ",
					idDetalleCaja));
         dValor = dValor -dValoresV;
			}
       this.lblEstado.Text = "Deposito: " +dValordeposito ;
				 //+ " / Valor pendiente de seleccionar " + dValor;
		}

		private void frmReporteValidaDeposito_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			decimal dDepositadoHastaahoratot = 0;
			 
			decimal dVerificar = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDepositos.Rows.All)
			{
				dVerificar = dVerificar + Convert.ToDecimal(dr.Cells["Valor"].Value) ;
			}
      dDepositadoHastaahoratot = dDepositadoHastaahora + dVerificar;

			if (dVerificar == dValor)
			{
				this.DialogResult = DialogResult.OK;
			}
			else
			{
				MessageBox.Show("El Total ingresado no pude ser diferente al deposito.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			
		//			if ( dVerificar < dValor )
		//			{
		//				if ( Convert.ToDecimal( dEfectivoCaja) == Convert.ToDecimal(dDepositadoHastaahoratot) )
		//				{
		//					this.DialogResult = DialogResult.OK;
		//				}
		//				else
		//				{
		//					MessageBox.Show("El Total ingresado no pude ser menor al deposito.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
		//					return;
		//				}
		//				
		//			}
		//			if ( dVerificar > dValor )
		//			{
		//				MessageBox.Show("El Total ingresado no pude ser mayor al deposito.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
		//				return;
		//			}
		//
		//			this.DialogResult = DialogResult.OK;
	}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
	
		}
	bool bActualiza = false;
		private void uGridDepositos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			
			if (bActualiza) return;

			if (e.Cell.Column.ToString() == "Valor")
			{	
				bActualiza = true;

				if (e.Cell.Row.Cells["Valor"].Value == System.DBNull.Value)
				{					
						e.Cell.Row.Cells["Valor"].Value = (decimal)e.Cell.OriginalValue;	
				}	
				if ((decimal ) e.Cell.Row.Cells["Valor"].Value < 0 )
				{					
					e.Cell.Row.Cells["Valor"].Value = (decimal)e.Cell.OriginalValue;	
				}
        decimal dSaldos = 0;
				//dSaldos = (decimal ) e.Cell.Row.Cells["Efectivo"].Value -  (decimal ) e.Cell.Row.Cells["Utilizado"].Value ;
				if ((decimal ) e.Cell.Row.Cells["Valor"].Value > (decimal ) e.Cell.Row.Cells["Disponible"].Value )
				{	
					MessageBox.Show("El valor ingresado no puede ser mayor al disponible.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cell.Row.Cells["Valor"].Value = (decimal)e.Cell.OriginalValue;	
				}
			
				bActualiza = false;
			}
		}

		private void uGridDepositos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Valor", 2);
		}

		private void uGridDepositos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridDepositos);						
		}
	}
}
