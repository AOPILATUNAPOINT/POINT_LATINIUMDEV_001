using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReembolsoViaticos.
	/// </summary>
	public class frmReembolsoViaticos : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDocumentos;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnCancelar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		public Infragistics.Win.UltraWinGrid.UltraGrid UgridDetalle;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReembolsoViaticos()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}
    bool bPantalla = false;// true ingresa nuevos gastos de viaje  false: ingresan los gastos el personal 
		bool bAdministrativoContabilidad = false;
		int idPersonal = 0;
		DateTime dtFechaSalida ;
		DateTime dtFechaRetorno ;
		int idGastosDeViaje = 0 ;
		int itipo = 0;
		public frmReembolsoViaticos(bool BPantalla)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			bPantalla = BPantalla;

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}
		public frmReembolsoViaticos( bool BAdministrativoContabilidad, int IdPersonal, DateTime DtFechaSalida, DateTime DtFechaRetorno, int IdGastosDeViaje, int Itipo)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			bAdministrativoContabilidad = BAdministrativoContabilidad;
			idPersonal = IdPersonal;
			dtFechaSalida = DtFechaSalida;
			dtFechaRetorno = DtFechaRetorno;
			idGastosDeViaje = IdGastosDeViaje;
			itipo = Itipo;

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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idReembolsoViaticos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Concepto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Administrativo");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Gerentes");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Diario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Activo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JefeDepartamental");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Administrativo", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Administrativo", 2, true);
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Gerentes", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Gerentes", 3, true);
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JefeDepartamental", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JefeDepartamental", 6, true);
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idReembolsoViaticos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Concepto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Administrativo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Gerentes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Diario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Activo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("JefeDepartamental");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReembolsoViaticos));
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idIngresoReembolsoViaticos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGastosViajes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idReembolsoViaticos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Definido");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Concepto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AdministrativoSolicitado");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("GerentesSolicitado");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Administrativo");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Gerentes");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotal");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Diario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SEL");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotalAcreditado");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SubTotal", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SubTotal", 9, true);
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Administrativo", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Administrativo", 7, true);
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SubTotalAcreditado", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SubTotalAcreditado", 12, true);
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idIngresoReembolsoViaticos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGastosViajes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idReembolsoViaticos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Definido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Concepto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AdministrativoSolicitado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("GerentesSolicitado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Administrativo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Gerentes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Diario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SEL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubTotalAcreditado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias");
			this.uGridDocumentos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.UgridDetalle = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.uGridDocumentos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.UgridDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridDocumentos
			// 
			this.uGridDocumentos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDocumentos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDocumentos.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridDocumentos.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn2.Header.VisiblePosition = 3;
			ultraGridColumn2.Width = 300;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn3.CellAppearance = appearance2;
			ultraGridColumn3.Format = " #,##0.00";
			ultraGridColumn3.Header.VisiblePosition = 4;
			ultraGridColumn3.Width = 95;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance3;
			ultraGridColumn4.Format = " #,##0.00";
			ultraGridColumn4.Header.VisiblePosition = 6;
			ultraGridColumn4.Width = 95;
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn6.Header.VisiblePosition = 1;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance4;
			ultraGridColumn7.Format = " #,##0.00";
			ultraGridColumn7.Header.Caption = "Jefe Departamental/Regional";
			ultraGridColumn7.Header.VisiblePosition = 5;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance5;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance6;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance7;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridDocumentos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridDocumentos.DisplayLayout.GroupByBox.Hidden = true;
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDocumentos.DisplayLayout.Override.ActiveRowAppearance = appearance8;
			this.uGridDocumentos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridDocumentos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridDocumentos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance9.BackColor = System.Drawing.Color.Transparent;
			this.uGridDocumentos.DisplayLayout.Override.CardAreaAppearance = appearance9;
			appearance10.ForeColor = System.Drawing.Color.Black;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDocumentos.DisplayLayout.Override.CellAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance11.FontData.BoldAsString = "True";
			appearance11.FontData.Name = "Arial";
			appearance11.FontData.SizeInPoints = 8F;
			appearance11.ForeColor = System.Drawing.Color.White;
			appearance11.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDocumentos.DisplayLayout.Override.HeaderAppearance = appearance11;
			this.uGridDocumentos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance12.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance12.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.RowAlternateAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.RowSelectorAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.SelectedRowAppearance = appearance14;
			this.uGridDocumentos.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridDocumentos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDocumentos.Location = new System.Drawing.Point(8, 8);
			this.uGridDocumentos.Name = "uGridDocumentos";
			this.uGridDocumentos.Size = new System.Drawing.Size(964, 315);
			this.uGridDocumentos.TabIndex = 766;
			this.uGridDocumentos.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridDocumentos_AfterRowInsert);
			this.uGridDocumentos.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridDocumentos_BeforeRowsDeleted);
			this.uGridDocumentos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridDocumentos_KeyDown);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(System.Decimal);
			ultraDataColumn4.DataType = typeof(System.Decimal);
			ultraDataColumn5.DataType = typeof(bool);
			ultraDataColumn6.DataType = typeof(bool);
			ultraDataColumn7.DataType = typeof(System.Decimal);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7});
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(8, 336);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(80, 23);
			this.btnGuardar.TabIndex = 769;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.ImageIndex = 62;
			this.btnSalir.Location = new System.Drawing.Point(184, 336);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(80, 23);
			this.btnSalir.TabIndex = 768;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnCancelar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(96, 336);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(80, 23);
			this.btnCancelar.TabIndex = 770;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
			// UgridDetalle
			// 
			this.UgridDetalle.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.UgridDetalle.Cursor = System.Windows.Forms.Cursors.Default;
			this.UgridDetalle.DataSource = this.ultraDataSource2;
			appearance15.BackColor = System.Drawing.Color.White;
			this.UgridDetalle.DisplayLayout.Appearance = appearance15;
			this.UgridDetalle.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance16;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.VisiblePosition = 5;
			ultraGridColumn12.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.VisiblePosition = 3;
			ultraGridColumn12.Width = 300;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance17;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.Caption = "$ Solicitud";
			ultraGridColumn13.Header.VisiblePosition = 6;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance18;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.Caption = "$ Solicitud";
			ultraGridColumn14.Header.VisiblePosition = 7;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance19;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.Caption = "$ Acreditado";
			ultraGridColumn15.Header.VisiblePosition = 11;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance20;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.VisiblePosition = 8;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance21.BackColor = System.Drawing.Color.Green;
			appearance21.BackColor2 = System.Drawing.Color.Green;
			appearance21.FontData.BoldAsString = "True";
			appearance21.FontData.ItalicAsString = "False";
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance21;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.VisiblePosition = 10;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn18.Header.VisiblePosition = 4;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 9;
			appearance22.BackColor = System.Drawing.Color.Green;
			appearance22.BackColor2 = System.Drawing.Color.Green;
			appearance22.FontData.BoldAsString = "True";
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance22;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.Caption = "$ SubTotal Acreditado";
			ultraGridColumn20.Header.VisiblePosition = 12;
			ultraGridColumn21.Header.VisiblePosition = 13;
			ultraGridColumn21.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21});
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance23;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance24;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance25;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings4,
																																															summarySettings5,
																																															summarySettings6});
			ultraGridBand2.SummaryFooterCaption = "";
			this.UgridDetalle.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.UgridDetalle.DisplayLayout.GroupByBox.Hidden = true;
			appearance26.ForeColor = System.Drawing.Color.Black;
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.UgridDetalle.DisplayLayout.Override.ActiveRowAppearance = appearance26;
			this.UgridDetalle.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.UgridDetalle.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.UgridDetalle.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance27.BackColor = System.Drawing.Color.Transparent;
			this.UgridDetalle.DisplayLayout.Override.CardAreaAppearance = appearance27;
			appearance28.ForeColor = System.Drawing.Color.Black;
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.UgridDetalle.DisplayLayout.Override.CellAppearance = appearance28;
			appearance29.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance29.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance29.FontData.BoldAsString = "True";
			appearance29.FontData.Name = "Arial";
			appearance29.FontData.SizeInPoints = 8F;
			appearance29.ForeColor = System.Drawing.Color.White;
			appearance29.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.UgridDetalle.DisplayLayout.Override.HeaderAppearance = appearance29;
			this.UgridDetalle.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance30.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance30.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance30.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.UgridDetalle.DisplayLayout.Override.RowAlternateAppearance = appearance30;
			appearance31.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance31.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance31.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.UgridDetalle.DisplayLayout.Override.RowSelectorAppearance = appearance31;
			appearance32.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance32.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance32.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.UgridDetalle.DisplayLayout.Override.SelectedRowAppearance = appearance32;
			this.UgridDetalle.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.UgridDetalle.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.UgridDetalle.Location = new System.Drawing.Point(8, 8);
			this.UgridDetalle.Name = "UgridDetalle";
			this.UgridDetalle.Size = new System.Drawing.Size(988, 315);
			this.UgridDetalle.TabIndex = 771;
			this.UgridDetalle.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.UgridDetalle_CellChange);
			this.UgridDetalle.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.UgridDetalle_AfterRowInsert);
			this.UgridDetalle.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.UgridDetalle_KeyPress);
			this.UgridDetalle.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.UgridDetalle_AfterCellUpdate);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(System.Decimal);
			ultraDataColumn13.DataType = typeof(System.Decimal);
			ultraDataColumn14.DataType = typeof(System.Decimal);
			ultraDataColumn15.DataType = typeof(System.Decimal);
			ultraDataColumn16.DataType = typeof(System.Decimal);
			ultraDataColumn17.DataType = typeof(System.Decimal);
			ultraDataColumn18.DataType = typeof(bool);
			ultraDataColumn19.DataType = typeof(bool);
			ultraDataColumn20.DataType = typeof(System.Decimal);
			ultraDataColumn21.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21});
			// 
			// frmReembolsoViaticos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1004, 366);
			this.ControlBox = false;
			this.Controls.Add(this.UgridDetalle);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.uGridDocumentos);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmReembolsoViaticos";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Viaticos Viajes a Provincia";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReembolsoViaticos_KeyDown);
			this.Load += new System.EventHandler(this.frmReembolsoViaticos_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridDocumentos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.UgridDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		int iGerAdm = 0;
		int idDias =0;

		private void uGridDocumentos_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idReembolsoViaticos"].Value = 0;
			e.Row.Cells["Concepto"].Value = "" ;	
			e.Row.Cells["Administrativo"].Value = 0.00m ;	
			e.Row.Cells["Gerentes"].Value = 0.00m ;	
			e.Row.Cells["Diario"].Value = false;	
			e.Row.Cells["Activo"].Value = false;	
			e.Row.Cells["JefeDepartamental"].Value = 0.00m ;	

		}

		private void uGridDocumentos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender,e,this.uGridDocumentos);
		}
		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}
		private void frmReembolsoViaticos_Load(object sender, System.EventArgs e)
		{
			
			if (bPantalla)
			{
				miAcceso = new Acceso(cdsSeteoF, "0220");
				if (!Funcion.Permiso("1266", cdsSeteoF))
				{				
					MessageBox.Show("No puede ingresar a Viaticos Gastos De Viaje ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					BeginInvoke(new MethodInvoker(UnloadMe));
					return;
				}

				this.UgridDetalle.Visible = false;
				this.uGridDocumentos.Visible = true;
				string sSQLPa = string.Format(" EXEC ConsultaReembolsoViaticos");
				FuncionesProcedimientos.dsGeneral(sSQLPa, this.uGridDocumentos);
			}
			else
			{
				idDias = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" SELECT DATEDIFF(DAY, '{0}', '{1}') +1", Convert.ToDateTime(dtFechaSalida).ToString("yyyyMMdd"),Convert.ToDateTime(dtFechaRetorno).ToString("yyyyMMdd") ));		
				iGerAdm = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" EXEC RetornaGerAdmin {0}  ", idPersonal));		

				this.UgridDetalle.Visible = true;
				this.uGridDocumentos.Visible = false;

				//				 idDias = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" SELECT DATEDIFF(DAY, '{0}', '{1}') +1", Convert.ToDateTime(dtFechaSalida).ToString("yyyyMMdd"),Convert.ToDateTime(dtFechaRetorno).ToString("yyyyMMdd") ));		
				if(!bAdministrativoContabilidad)
				{
					#region region Gastos vacios
					if ((int) idGastosDeViaje  == 0)
					{
						string sSQLPa = string.Format(" EXEC ConsultaIngresoReembolsoViaticos  {0}, {1},  {2}, {3}  ", bAdministrativoContabilidad, idPersonal, idDias, itipo );
						FuncionesProcedimientos.dsGeneral(sSQLPa, this.UgridDetalle);

						iGerAdm = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" EXEC RetornaGerAdmin {0}  ", idPersonal));		

						if (iGerAdm == 0 && !bAdministrativoContabilidad)
						{
							this.UgridDetalle.DisplayLayout.Bands[0].Columns["GerentesSolicitado"].Hidden = true;
							this.UgridDetalle.DisplayLayout.Bands[0].Columns["Administrativo"].Hidden = true;
							this.UgridDetalle.DisplayLayout.Bands[0].Columns["Gerentes"].Hidden = true;
							this.UgridDetalle.DisplayLayout.Bands[0].Columns["SubTotalAcreditado"].Hidden = true;
						}

					
						if (iGerAdm == 1 && !bAdministrativoContabilidad)
						{
							this.UgridDetalle.DisplayLayout.Bands[0].Columns["AdministrativoSolicitado"].Hidden = true;
							this.UgridDetalle.DisplayLayout.Bands[0].Columns["Administrativo"].Hidden = true;
							this.UgridDetalle.DisplayLayout.Bands[0].Columns["Gerentes"].Hidden = true;
							this.UgridDetalle.DisplayLayout.Bands[0].Columns["SubTotalAcreditado"].Hidden = true;
						}
						#endregion region Gastos vacios
					}
					else
					{
						string sSQLPa = string.Format(" EXEC ConsultaDetalleIngresoReembolsoViaticos   {0}, {1}, {2} ", idGastosDeViaje, idDias, iGerAdm);
						FuncionesProcedimientos.dsGeneral(sSQLPa, this.UgridDetalle);

						
						if (iGerAdm == 0 && !bAdministrativoContabilidad)
						{
							this.UgridDetalle.DisplayLayout.Bands[0].Columns["GerentesSolicitado"].Hidden = true;
							this.UgridDetalle.DisplayLayout.Bands[0].Columns["Administrativo"].Hidden = true;
							this.UgridDetalle.DisplayLayout.Bands[0].Columns["Gerentes"].Hidden = true;
							this.UgridDetalle.DisplayLayout.Bands[0].Columns["SubTotalAcreditado"].Hidden = true;
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.UgridDetalle.Rows.All)
							{	
								if ((bool)dr.Cells["SEL"].Value )
								{
									if ((bool) dr.Cells["Diario"].Value  )
									{
										dr.Cells["SubTotal"].Value = 0 ;
										dr.Cells["SubTotal"].Value = Convert.ToDecimal (dr.Cells["AdministrativoSolicitado"].Value)  * Convert.ToDecimal ( dr.Cells["Dias"].Value) ;
									}
								}
							}
						}
					
						if (iGerAdm == 1 && !bAdministrativoContabilidad)
						{
							this.UgridDetalle.DisplayLayout.Bands[0].Columns["AdministrativoSolicitado"].Hidden = true;
							this.UgridDetalle.DisplayLayout.Bands[0].Columns["Administrativo"].Hidden = true;
							this.UgridDetalle.DisplayLayout.Bands[0].Columns["Gerentes"].Hidden = true;
							this.UgridDetalle.DisplayLayout.Bands[0].Columns["SubTotalAcreditado"].Hidden = true;
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.UgridDetalle.Rows.All)
							{	
								if ((bool)dr.Cells["SEL"].Value )
								{
									if ((bool) dr.Cells["Diario"].Value )
									{
										dr.Cells["SubTotal"].Value = 0 ;
										dr.Cells["SubTotal"].Value = Convert.ToDecimal (dr.Cells["GerentesSolicitado"].Value)  * Convert.ToDecimal ( dr.Cells["Dias"].Value) ;
									}
								}
							}
						}
					}
				}
				else
				{
					string sSQLPa = string.Format(" EXEC ConsultaDetalleIngresoReembolsoViaticos   {0}, {1}, {2} ", idGastosDeViaje, idDias, iGerAdm);
					FuncionesProcedimientos.dsGeneral(sSQLPa, this.UgridDetalle);
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.UgridDetalle.Rows.All)
					{
						dr.Cells["Administrativo"].Value = Convert.ToDecimal( dr.Cells["GerentesSolicitado"].Value)+ Convert.ToDecimal( dr.Cells["AdministrativoSolicitado"].Value);
						dr.Cells["SubTotalAcreditado"].Value = Convert.ToDecimal( dr.Cells["SubTotal"].Value);
					}
					if (iGerAdm == 0 && bAdministrativoContabilidad)
					{
						this.UgridDetalle.DisplayLayout.Bands[0].Columns["GerentesSolicitado"].Hidden = true;
						this.UgridDetalle.DisplayLayout.Bands[0].Columns["AdministrativoSolicitado"].Hidden = true;
						this.UgridDetalle.DisplayLayout.Bands[0].Columns["SEL"].Hidden = true;
						this.UgridDetalle.DisplayLayout.Bands[0].Columns["SubTotal"].Hidden = true;
						this.UgridDetalle.DisplayLayout.Bands[0].Columns["SEL"].CellActivation =  Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
						this.UgridDetalle.DisplayLayout.Bands[0].Columns["SubTotal"].CellActivation =  Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
						this.UgridDetalle.DisplayLayout.Bands[0].Columns["SubTotalAcreditado"].CellActivation =  Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
						this.UgridDetalle.DisplayLayout.Bands[0].Columns["GerentesSolicitado"].CellActivation =  Infragistics.Win.UltraWinGrid.Activation.NoEdit;
						this.UgridDetalle.DisplayLayout.Bands[0].Columns["AdministrativoSolicitado"].CellActivation =  Infragistics.Win.UltraWinGrid.Activation.NoEdit;
						this.UgridDetalle.DisplayLayout.Bands[0].Columns["Administrativo"].CellActivation =  Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
						this.UgridDetalle.DisplayLayout.Bands[0].Columns["Gerentes"].CellActivation =  Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
						this.UgridDetalle.DisplayLayout.Bands[0].Columns["Gerentes"].Hidden = true;
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.UgridDetalle.Rows.All)
						{	
							if ((bool)dr.Cells["SEL"].Value )
							{
								dr.Cells["SubTotal"].Value = Convert.ToDecimal (dr.Cells["AdministrativoSolicitado"].Value)  * Convert.ToDecimal ( dr.Cells["Dias"].Value) ;
							}
						}
					}
					
					if (iGerAdm == 1 && bAdministrativoContabilidad)
					{
						this.UgridDetalle.DisplayLayout.Bands[0].Columns["GerentesSolicitado"].Hidden = true;
						this.UgridDetalle.DisplayLayout.Bands[0].Columns["AdministrativoSolicitado"].Hidden = true;
						this.UgridDetalle.DisplayLayout.Bands[0].Columns["SEL"].Hidden = true;
						this.UgridDetalle.DisplayLayout.Bands[0].Columns["SubTotal"].Hidden = true;
						this.UgridDetalle.DisplayLayout.Bands[0].Columns["GerentesSolicitado"].CellActivation =  Infragistics.Win.UltraWinGrid.Activation.NoEdit;
						this.UgridDetalle.DisplayLayout.Bands[0].Columns["AdministrativoSolicitado"].CellActivation =  Infragistics.Win.UltraWinGrid.Activation.NoEdit;
						this.UgridDetalle.DisplayLayout.Bands[0].Columns["SEL"].CellActivation =  Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
						this.UgridDetalle.DisplayLayout.Bands[0].Columns["SubTotal"].CellActivation =  Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
						this.UgridDetalle.DisplayLayout.Bands[0].Columns["Administrativo"].CellActivation =  Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
						this.UgridDetalle.DisplayLayout.Bands[0].Columns["Gerentes"].CellActivation =  Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
								this.UgridDetalle.DisplayLayout.Bands[0].Columns["SubTotalAcreditado"].CellActivation =  Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
						this.UgridDetalle.DisplayLayout.Bands[0].Columns["Gerentes"].Hidden = true;
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.UgridDetalle.Rows.All)
						{	
							if ((bool)dr.Cells["SEL"].Value )
							{
								dr.Cells["SubTotal"].Value = Convert.ToDecimal (dr.Cells["GerentesSolicitado"].Value)  * Convert.ToDecimal ( dr.Cells["Dias"].Value) ;
							}
						}
					}

				}
			}
			if(idPersonal > 0) this.btnCancelar.Visible = false;
			if( idGastosDeViaje > 0 && !bAdministrativoContabilidad)
			{
				this.btnGuardar.Visible= false;
				this.UgridDetalle.DisplayLayout.Bands[0].Columns["SEL"].CellActivation =  Infragistics.Win.UltraWinGrid.Activation.NoEdit;
				this.UgridDetalle.DisplayLayout.Bands[0].Columns["SubTotal"].Hidden = true;
				this.UgridDetalle.DisplayLayout.Bands[0].Columns["SubTotalAcreditado"].Hidden = false;
				this.UgridDetalle.DisplayLayout.Bands[0].Columns["SubTotalAcreditado"].CellActivation =  Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
						
				FuncionesProcedimientos.EstadoGrids(false,this.UgridDetalle);
			}

		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			if (bPantalla)
			{
				string sSQLPa = string.Format(" EXEC ConsultaReembolsoViaticos");
				FuncionesProcedimientos.dsGeneral(sSQLPa, this.uGridDocumentos);
			}
			else
			{
				if (!bAdministrativoContabilidad)
				{
					int idDias = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" SELECT DATEDIFF(DAY, '{0}', '{1}')+1", Convert.ToDateTime(dtFechaSalida).ToString("yyyyMMdd"),Convert.ToDateTime(dtFechaRetorno).ToString("yyyyMMdd") ));		


					string sSQLPa = string.Format(" EXEC ConsultaIngresoReembolsoViaticos {0}, {1},  {2}  ", bAdministrativoContabilidad, idPersonal, idDias );
					FuncionesProcedimientos.dsGeneral(sSQLPa, this.UgridDetalle);
				}
				else
				{
					int idDias = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" SELECT DATEDIFF(DAY, '{0}', '{1}')+1", Convert.ToDateTime(dtFechaSalida).ToString("yyyyMMdd"),Convert.ToDateTime(dtFechaRetorno).ToString("yyyyMMdd") ));		

					string sSQLPa = string.Format(" EXEC ConsultaDetalleIngresoReembolsoViaticos   {0}, {1}, {2} ", idGastosDeViaje, idDias, iGerAdm);
					FuncionesProcedimientos.dsGeneral(sSQLPa, this.UgridDetalle);
				}

			}
		
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if (bPantalla)
			{
				#region Guarda Gastos Nuevos 
				if (bPantalla)
				{

					int iValores = 0;
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)
					{	
						if (dr.Cells["Concepto"].Value.ToString().Length == 0)
						{
							MessageBox.Show("Escriba el Concepto ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.uGridDocumentos.ActiveRow = dr;
							this.uGridDocumentos.ActiveRow.Selected = true;
							return;
						}
						if (dr.Cells["Concepto"].Value.ToString().Length < 10)
						{
							MessageBox.Show("El Concepto debe tener minimo 5 caracteres", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.uGridDocumentos.ActiveRow = dr;
							this.uGridDocumentos.ActiveRow.Selected = true;
							return;
						}

						if ( (decimal)dr.Cells["Administrativo"].Value == 0 || (decimal)dr.Cells["Gerentes"].Value == 0 || (decimal)dr.Cells["JefeDepartamental"].Value == 0  )
						{
							iValores = iValores + 1;
						}

					}

					if (DialogResult.Yes == MessageBox.Show(string.Format(" Desea Guardar Los Cambios " , 0), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
					 
					{
						using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
						{
							oConexion.Open();

							SqlTransaction oTransaction = null;
							SqlCommand oCmdActualiza = oConexion.CreateCommand();
							oCmdActualiza.Connection = oConexion;
							oCmdActualiza.CommandType = CommandType.Text;
							oCmdActualiza.CommandTimeout = 0;

							try
							{
								oTransaction = oConexion.BeginTransaction();
								oCmdActualiza.Transaction = oTransaction;

								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)
								{
									string  sSQLDet = string.Format(" Exec  GuardaReembolsoViaticos {0}, '{1}', {2}, {3}, {4}, {5} ,{6} ", 
										(int)dr.Cells["idReembolsoViaticos"].Value , Convert.ToString( dr.Cells["Concepto"].Value.ToString()),
										Convert.ToDecimal( dr.Cells["Administrativo"].Value.ToString()), Convert.ToDecimal( dr.Cells["Gerentes"].Value.ToString()),
										Convert.ToBoolean( dr.Cells["Diario"].Value.ToString()), Convert.ToDecimal( dr.Cells["JefeDepartamental"].Value.ToString()),
										Convert.ToBoolean( dr.Cells["Activo"].Value));
									oCmdActualiza.CommandText = sSQLDet;
									oCmdActualiza.ExecuteNonQuery();	
								}
								oTransaction.Commit();

								MessageBox.Show("Registro Guardado Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
								string sSQLPa = string.Format(" EXEC ConsultaReembolsoViaticos");
								FuncionesProcedimientos.dsGeneral(sSQLPa, this.uGridDocumentos);
					
							}
							catch(Exception ex)
							{
								try
								{ 
									oTransaction.Rollback();

									MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
								}
								catch (Exception ex2)
								{
									MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
								}
							}
							finally
							{
								oConexion.Close();
							}
						}
						this.DialogResult = DialogResult.OK;
					}

				}
				#endregion Guarda Gastos Nuevos 
			}
			else
			{
				this.DialogResult = DialogResult.OK;
			}
			
		}

		private void frmReembolsoViaticos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();		
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			if (bPantalla)
			{
				this.Close();		
			}
			else
			{
				if(!bAdministrativoContabilidad)
				{
					if ( idGastosDeViaje > 0)
					{
						this.DialogResult = DialogResult.OK;
					}
					else
					{
						this.DialogResult = DialogResult.No;

					}
				}
				else
				{
          this.DialogResult = DialogResult.No;
				}
			}
		}

		private void uGridDocumentos_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (bPantalla)
			{
				#region Elimina
				if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de borrar el Registro?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
				{				
					for (int i=0; i < e.Rows.Length; i++)
					{	
						if ((int)e.Rows[i].Cells["idReembolsoViaticos"].Value > 0)
						{
							string sSQL = string.Format("Delete From ReembolsoViaticos where idReembolsoViaticos = {0} ", (int)e.Rows[i].Cells["idReembolsoViaticos"].Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

							e.DisplayPromptMsg = false;
						}
						else e.DisplayPromptMsg = false;
					}					
				}
				else e.Cancel = true;
				#endregion Elimina
			}
		}

		private void UgridDetalle_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idIngresoReembolsoViaticos"].Value = 0;
			e.Row.Cells["idGastosViajes"].Value = 0;	
			e.Row.Cells["idReembolsoViaticos"].Value = 0;	
			e.Row.Cells["Concepto"].Value = 0.00m ;	
			e.Row.Cells["Definido"].Value = 0.00m ;
			e.Row.Cells["AdministrativoSolicitado"].Value = 0.00m ;	
			e.Row.Cells["GerenteSolicitado"].Value = 0.00m ;	
			e.Row.Cells["Administrativo"].Value = 0.00m ;	
			e.Row.Cells["Gerentes"].Value = 0.00m ;	
			e.Row.Cells["Total"].Value = 0.00m ;	
			e.Row.Cells["SubTotal"].Value = 0.00m ;	
			e.Row.Cells["diario"].Value = 0;	
			e.Row.Cells["SEL"].Value = 0;	

		}

		private void UgridDetalle_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if ( UgridDetalle.ActiveCell.Column.Key.ToString().ToUpper() == "SOLICITADO" )
			{
				Funcion.CamposNumericos(sender, e);
			}
			if ( UgridDetalle.ActiveCell.Column.Key.ToString().ToUpper() == "ADMINISTRATIVO" )
			{
				Funcion.CamposNumericos(sender, e);
			}
		}

		private void UgridDetalle_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			
		}

		private void UgridDetalle_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if(!bAdministrativoContabilidad)
			{
				if (iGerAdm == 0)
				{
					#region Pantalla Ingreso Administrativo
					if (e.Cell.Column.ToString() == "SEL")
					{	
						this.UgridDetalle.UpdateData();
						if (e.Cell.Row.Cells["AdministrativoSolicitado"].Value  == System.DBNull.Value) 	e.Cell.Row.Cells["AdministrativoSolicitado"].Value = (decimal)e.Cell.OriginalValue;
					
					
						if ((bool)e.Cell.Row.Cells["SEL"].Value == true)
						{
							if (((bool)e.Cell.Row.Cells["diario"].Value == true))
							{
								e.Cell.Row.Cells["AdministrativoSolicitado"].Value = (decimal)e.Cell.Row.Cells["Definido"].Value ;
								e.Cell.Row.Cells["SubTotal"].Value = (decimal)e.Cell.Row.Cells["AdministrativoSolicitado"].Value  * Convert.ToDecimal ( e.Cell.Row.Cells["Dias"].Value) ;
							}
							else
							{
								e.Cell.Row.Cells["AdministrativoSolicitado"].Value = (decimal)e.Cell.Row.Cells["Definido"].Value ;
								e.Cell.Row.Cells["SubTotal"].Value = (decimal)e.Cell.Row.Cells["AdministrativoSolicitado"].Value ;
							}
						}
						else
						{
							e.Cell.Row.Cells["AdministrativoSolicitado"].Value = 0.00m ;	
							e.Cell.Row.Cells["SubTotal"].Value = 0.00m ;	
						}
					}
      
					if (e.Cell.Column.ToString() == "AdministrativoSolicitado")
					{		
						this.UgridDetalle.UpdateData();
						if (e.Cell.Row.Cells["AdministrativoSolicitado"].Value  == System.DBNull.Value) 	e.Cell.Row.Cells["AdministrativoSolicitado"].Value = (decimal)e.Cell.OriginalValue;
			
						if ((decimal)e.Cell.Row.Cells["Definido"].Value > 0)
						{
        
							if ((decimal)e.Cell.Row.Cells["AdministrativoSolicitado"].Value > (decimal)e.Cell.Row.Cells["Definido"].Value )
							{
								MessageBox.Show("El valor Solicitado no puede ser Mayor al Valor Definido ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

								e.Cell.Row.Cells["AdministrativoSolicitado"].Value = (decimal)e.Cell.OriginalValue;
					
							}
					
							if ((bool)e.Cell.Row.Cells["SEL"].Value == true)
							{
								if (((bool)e.Cell.Row.Cells["diario"].Value == true))
								{
									e.Cell.Row.Cells["SubTotal"].Value = (decimal)e.Cell.Row.Cells["AdministrativoSolicitado"].Value  * Convert.ToDecimal ( e.Cell.Row.Cells["Dias"].Value) ;
								}
								else
								{
//									e.Cell.Row.Cells["AdministrativoSolicitado"].Value = (decimal)e.Cell.Row.Cells["Definido"].Value ;
									e.Cell.Row.Cells["SubTotal"].Value = (decimal)e.Cell.Row.Cells["AdministrativoSolicitado"].Value ;
								}
							}
							else
							{
								e.Cell.Row.Cells["AdministrativoSolicitado"].Value = 0.00m ;	
								e.Cell.Row.Cells["SubTotal"].Value = 0.00m ;	
							}
						}
						else
						{
							if ((bool)e.Cell.Row.Cells["SEL"].Value == true)
							{
								if (((bool)e.Cell.Row.Cells["diario"].Value == true))
								{
									e.Cell.Row.Cells["SubTotal"].Value = (decimal)e.Cell.Row.Cells["AdministrativoSolicitado"].Value  * Convert.ToDecimal ( e.Cell.Row.Cells["Dias"].Value) ;
								}
								else
								{
									e.Cell.Row.Cells["SubTotal"].Value = (decimal)e.Cell.Row.Cells["AdministrativoSolicitado"].Value ;
								}
							}
							else
							{
								e.Cell.Row.Cells["AdministrativoSolicitado"].Value = 0.00m ;	
								e.Cell.Row.Cells["SubTotal"].Value = 0.00m ;	
							}
						}
				
					}
					#endregion Pantalla Ingreso Administrativo
				}
				else
				{
					#region Pantalla Ingreso Administrativo
					if (e.Cell.Column.ToString() == "SEL")
					{	
						this.UgridDetalle.UpdateData();
						if (e.Cell.Row.Cells["GerentesSolicitado"].Value  == System.DBNull.Value) 	e.Cell.Row.Cells["GerentesSolicitado"].Value = (decimal)e.Cell.OriginalValue;
					
					
						if ((bool)e.Cell.Row.Cells["SEL"].Value == true)
						{
							if (((bool)e.Cell.Row.Cells["diario"].Value == true))
							{
								e.Cell.Row.Cells["GerentesSolicitado"].Value = (decimal)e.Cell.Row.Cells["Definido"].Value ;
								e.Cell.Row.Cells["SubTotal"].Value = (decimal)e.Cell.Row.Cells["GerentesSolicitado"].Value  * Convert.ToDecimal ( e.Cell.Row.Cells["Dias"].Value) ;
							}
							else
							{
//								e.Cell.Row.Cells["GerentesSolicitado"].Value = (decimal)e.Cell.Row.Cells["Definido"].Value ;
            e.Cell.Row.Cells["SubTotal"].Value = (decimal)e.Cell.Row.Cells["GerentesSolicitado"].Value ;
							}
						}
						else
						{
							e.Cell.Row.Cells["GerentesSolicitado"].Value = 0.00m ;	
							e.Cell.Row.Cells["SubTotal"].Value = 0.00m ;	
						}
					}
      
					if (e.Cell.Column.ToString() == "GerentesSolicitado")
					{		
						this.UgridDetalle.UpdateData();
						if (e.Cell.Row.Cells["GerentesSolicitado"].Value  == System.DBNull.Value) 	e.Cell.Row.Cells["GerentesSolicitado"].Value = (decimal)e.Cell.OriginalValue;
			
						if ((decimal)e.Cell.Row.Cells["Definido"].Value > 0)
						{
        
							if ((decimal)e.Cell.Row.Cells["GerentesSolicitado"].Value > (decimal)e.Cell.Row.Cells["Definido"].Value )
							{
								MessageBox.Show("El valor Solicitado no puede ser Mayor al Valor Definido ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

								e.Cell.Row.Cells["GerentesSolicitado"].Value = (decimal)e.Cell.OriginalValue;
					
							}
					
							if ((bool)e.Cell.Row.Cells["SEL"].Value == true)
							{
								if (((bool)e.Cell.Row.Cells["diario"].Value == true))
								{
									e.Cell.Row.Cells["GerentesSolicitado"].Value = (decimal)e.Cell.Row.Cells["Definido"].Value ;
									e.Cell.Row.Cells["SubTotal"].Value = (decimal)e.Cell.Row.Cells["GerentesSolicitado"].Value  * Convert.ToDecimal ( e.Cell.Row.Cells["Dias"].Value) ;
								}
								else
								{
//									e.Cell.Row.Cells["GerentesSolicitado"].Value = (decimal)e.Cell.Row.Cells["Definido"].Value ;
  								e.Cell.Row.Cells["SubTotal"].Value = (decimal)e.Cell.Row.Cells["GerentesSolicitado"].Value ;
								}
							}
							else
							{
								e.Cell.Row.Cells["GerentesSolicitado"].Value = 0.00m ;	
								e.Cell.Row.Cells["SubTotal"].Value = 0.00m ;	
							}
						}
						else
						{
							if ((bool)e.Cell.Row.Cells["SEL"].Value == true)
							{
								if (((bool)e.Cell.Row.Cells["diario"].Value == true))
								{
									e.Cell.Row.Cells["SubTotal"].Value = (decimal)e.Cell.Row.Cells["GerentesSolicitado"].Value  * Convert.ToDecimal ( e.Cell.Row.Cells["Dias"].Value) ;
								}
								else
								{
									e.Cell.Row.Cells["SubTotal"].Value = (decimal)e.Cell.Row.Cells["GerentesSolicitado"].Value ;
								}
							}
							else
							{
								e.Cell.Row.Cells["GerentesSolicitado"].Value = 0.00m ;	
								e.Cell.Row.Cells["SubTotal"].Value = 0.00m ;	
							}
						}
				
					}
					#endregion Pantalla Ingreso Administrativo

				}
			}
			else
			{
				
					#region Pantalla Ingreso Administrativo
					
					if (e.Cell.Column.ToString() == "Administrativo")
					{		
						this.UgridDetalle.UpdateData();
						if (e.Cell.Row.Cells["Administrativo"].Value  == System.DBNull.Value) 	e.Cell.Row.Cells["Administrativo"].Value = (decimal)e.Cell.OriginalValue;
			
						if ((decimal)e.Cell.Row.Cells["Definido"].Value > 0)
						{
        
//							if ((decimal)e.Cell.Row.Cells["Administrativo"].Value > (decimal)e.Cell.Row.Cells["Definido"].Value )
//							{
//								MessageBox.Show("El valor Solicitado no puede ser Mayor al Valor Definido ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//
//								e.Cell.Row.Cells["Administrativo"].Value = (decimal)e.Cell.OriginalValue;
//					
//							}
							if (((bool)e.Cell.Row.Cells["diario"].Value == true))
							{
								e.Cell.Row.Cells["SubTotalAcreditado"].Value = (decimal)e.Cell.Row.Cells["Administrativo"].Value  * Convert.ToDecimal ( e.Cell.Row.Cells["Dias"].Value) ;
							}
							else
							{
								e.Cell.Row.Cells["SubTotalAcreditado"].Value = (decimal)e.Cell.Row.Cells["Administrativo"].Value ;
							}
						
						}
						else
						{
							if (((bool)e.Cell.Row.Cells["diario"].Value == true))
							{
								e.Cell.Row.Cells["SubTotalAcreditado"].Value = (decimal)e.Cell.Row.Cells["Administrativo"].Value  * Convert.ToDecimal ( e.Cell.Row.Cells["Dias"].Value) ;
							}
							else
							{
								e.Cell.Row.Cells["SubTotalAcreditado"].Value = (decimal)e.Cell.Row.Cells["Administrativo"].Value ;
							}
						
						}
					}
					#endregion Pantalla Ingreso Administrativo
				
	
			}

			}
		}
	}

