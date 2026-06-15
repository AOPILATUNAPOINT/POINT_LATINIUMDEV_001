using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Text;
using System.Threading;
using C1.Data;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmFactDentalis.
	/// </summary>
	public class frmFactDentalis : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion;
		private System.Windows.Forms.Button btnImportar;
		private System.Windows.Forms.Button btnExcel;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmFactDentalis()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRegFactDentalis");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Establecimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoIdentificacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Rut");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Paciente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TELEFONO");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DIRECCION");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Presupuesto");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ConvenioPresupuesto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EspecialidadPresupuesto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ConvenioPac");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Institucion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Diagnosticador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comprobante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Caja");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BancoTarjeta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("LOTE");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TRASSACCION");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TURNO");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FACTURA");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumDoc");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vcto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dcto_ppto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRegFactDentalis");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sucursal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Establecimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoIdentificacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Rut");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Paciente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TELEFONO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DIRECCION");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Presupuesto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ConvenioPresupuesto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EspecialidadPresupuesto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ConvenioPac");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Institucion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Diagnosticador");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Comprobante");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Caja");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("BancoTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("LOTE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TRASSACCION");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TURNO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FACTURA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumDoc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vcto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dcto_ppto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorFactura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmFactDentalis));
			this.uGridInformacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnImportar = new System.Windows.Forms.Button();
			this.btnExcel = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
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
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 25;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 90;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn5.Header.Caption = "Tipo Identificacion";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 120;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 90;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 250;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 250;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance2;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 80;
			ultraGridColumn11.Header.Caption = "Convenio Presupuesto";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 130;
			ultraGridColumn12.Header.Caption = "Especialidad Presupuesto";
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 140;
			ultraGridColumn13.Header.Caption = "Convenio Pac";
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Width = 95;
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn18.Header.Caption = "Banco / Tarjeta";
			ultraGridColumn18.Header.VisiblePosition = 17;
			ultraGridColumn18.Width = 150;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance3;
			ultraGridColumn19.Header.VisiblePosition = 18;
			ultraGridColumn19.Width = 60;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance4;
			ultraGridColumn20.Header.VisiblePosition = 19;
			ultraGridColumn21.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn21.CellAppearance = appearance5;
			ultraGridColumn21.Header.VisiblePosition = 20;
			ultraGridColumn21.Width = 50;
			ultraGridColumn22.Header.VisiblePosition = 21;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance6;
			ultraGridColumn23.Header.Caption = "N° Doc";
			ultraGridColumn23.Header.VisiblePosition = 22;
			ultraGridColumn23.Width = 80;
			ultraGridColumn24.Header.VisiblePosition = 23;
			ultraGridColumn25.Header.VisiblePosition = 24;
			ultraGridColumn26.Header.Caption = "Valor Factura";
			ultraGridColumn26.Header.VisiblePosition = 25;
			ultraGridColumn27.Header.VisiblePosition = 26;
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
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27});
			ultraGridBand1.Header.Caption = "Periodo A";
			ultraGridBand1.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			ultraGridBand1.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridInformacion.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridInformacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridInformacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridInformacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance7.BackColor = System.Drawing.Color.Transparent;
			this.uGridInformacion.DisplayLayout.Override.CardAreaAppearance = appearance7;
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInformacion.DisplayLayout.Override.CellAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance9.FontData.BoldAsString = "True";
			appearance9.FontData.Name = "Arial";
			appearance9.FontData.SizeInPoints = 8F;
			appearance9.ForeColor = System.Drawing.Color.White;
			appearance9.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInformacion.DisplayLayout.Override.HeaderAppearance = appearance9;
			this.uGridInformacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowSelectorAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.SelectedRowAppearance = appearance11;
			this.uGridInformacion.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInformacion.Location = new System.Drawing.Point(8, 80);
			this.uGridInformacion.Name = "uGridInformacion";
			this.uGridInformacion.Size = new System.Drawing.Size(1248, 424);
			this.uGridInformacion.TabIndex = 768;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(System.DateTime);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn19.DataType = typeof(int);
			ultraDataColumn20.DataType = typeof(int);
			ultraDataColumn23.DataType = typeof(int);
			ultraDataColumn24.DataType = typeof(System.DateTime);
			ultraDataColumn26.DataType = typeof(System.Decimal);
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
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21,
																																 ultraDataColumn22,
																																 ultraDataColumn23,
																																 ultraDataColumn24,
																																 ultraDataColumn25,
																																 ultraDataColumn26,
																																 ultraDataColumn27});
			// 
			// btnImportar
			// 
			this.btnImportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnImportar.CausesValidation = false;
			this.btnImportar.Image = ((System.Drawing.Image)(resources.GetObject("btnImportar.Image")));
			this.btnImportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImportar.Location = new System.Drawing.Point(1064, 40);
			this.btnImportar.Name = "btnImportar";
			this.btnImportar.Size = new System.Drawing.Size(88, 24);
			this.btnImportar.TabIndex = 776;
			this.btnImportar.Text = "Importar";
			this.btnImportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImportar.Click += new System.EventHandler(this.btnImportar_Click);
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(1160, 40);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(96, 24);
			this.btnExcel.TabIndex = 775;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
			// frmFactDentalis
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1264, 526);
			this.Controls.Add(this.btnImportar);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.uGridInformacion);
			this.Name = "frmFactDentalis";
			this.Text = "frmFactDentalis";
			this.Load += new System.EventHandler(this.frmFactDentalis_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnImportar_Click(object sender, System.EventArgs e)
		{
			#region Variables

			int iCont = 0;
			int IdFactDeantalis = 0;
			string sSucursal = "";
			string sEstablecimiento = "";
			string sFecha= "";
			string sTipoIdentificacion = "";
			string sRut = "";
			string sPaciente = "";
			string sTELEFONO = "";
			string sDIRECCION = "";
			int iPresupuesto = 0;
			string sConvenioPresupuesto = "";
			string sEspecialidadPresupuesto = "";
			string sConvenioPac = "";
			string sInstitucion = "";
			string sDiagnosticador = "";
			string sComprobante = "";
			string sCaja = "";
			string sBancoTarjeta = ""; 
			int iLOTE = 0;
			int iTRANSSACCION = 0;
			string sTURNO = ""; 
			string sFACTURA  = "";
			int iNumDoc = 0;
			string sVcto = "";
			string sDcto_ppto = ""; 
			decimal dValorFactura =  0.00m;
			string sEstado  = "";

			#endregion Variables

			#region Sube archivo
			try
			{
				OpenFileDialog sArchivo = new OpenFileDialog();

				sArchivo.Filter = "Microsoft Excel (*.xls)|*.xls";

				if (sArchivo.ShowDialog() == System.Windows.Forms.DialogResult.OK && sArchivo.FileName.Length > 0)
				{
					this.Cursor = Cursors.WaitCursor;
										
					foreach (DataRow row in FuncionesProcedimientos.ImportaDeExcel(sArchivo.FileName, "SELECT * FROM [DATOS$]", "DATOS").Tables[0].Rows)
					{						
						if (row["Sucursal"] != System.DBNull.Value) sSucursal = row["Sucursal"].ToString();
						if (row["Establecimiento"] != System.DBNull.Value) sEstablecimiento = row["Establecimiento"].ToString();
						if (row["Fecha"] != System.DBNull.Value) sFecha = row["Fecha"].ToString();
						if (row["TipoIdentificacion"] != System.DBNull.Value) sTipoIdentificacion = row["TipoIdentificacion"].ToString();
						if (row["Rut"] != System.DBNull.Value) sRut = row["Rut"].ToString();
						if (row["Paciente"] != System.DBNull.Value) sPaciente = row["Paciente"].ToString();
						if (row["TELEFONO"] != System.DBNull.Value) sTELEFONO = row["TELEFONO"].ToString();
						if (row["DIRECCION"] != System.DBNull.Value) sDIRECCION = row["DIRECCION"].ToString();
						if (row["Presupuesto"] != System.DBNull.Value) iPresupuesto = int.Parse(row["Presupuesto"].ToString());
						if (row["ConvenioPresupuesto"] != System.DBNull.Value) sConvenioPresupuesto = row["ConvenioPresupuesto"].ToString();
						if (row["EspecialidadPresupuesto"] != System.DBNull.Value) sEspecialidadPresupuesto = row["EspecialidadPresupuesto"].ToString();
						if (row["ConvenioPac"] != System.DBNull.Value) sConvenioPac = row["ConvenioPac"].ToString();
						if (row["Institucion"] != System.DBNull.Value) sInstitucion = row["Institucion"].ToString();
						if (row["Diagnosticador"] != System.DBNull.Value) sDiagnosticador = row["Diagnosticador"].ToString();
						if (row["Comprobante"] != System.DBNull.Value) sComprobante = row["Comprobante"].ToString();
						if (row["Caja"] != System.DBNull.Value) sCaja = row["Caja"].ToString();
						if (row["BancoTarjeta"] != System.DBNull.Value) sBancoTarjeta = row["BancoTarjeta"].ToString();
						if (row["LOTE"] != System.DBNull.Value) iLOTE = int.Parse(row["LOTE"].ToString());
						if (row["TRANSSACCION"] != System.DBNull.Value) iTRANSSACCION = int.Parse(row["TRANSSACCION"].ToString());
						if (row["TURNO"] != System.DBNull.Value) sTURNO = row["TURNO"].ToString();
						if (row["FACTURA"] != System.DBNull.Value) sFACTURA = row["FACTURA"].ToString();
						if (row["NumDoc"] != System.DBNull.Value) iNumDoc = int.Parse(row["NumDoc"].ToString());
						if (row["Vcto"] != System.DBNull.Value) sVcto = row["Vcto"].ToString();
						if (row["Dcto_ppto"] != System.DBNull.Value) sDcto_ppto = row["Dcto_ppto"].ToString();
						if (row["ValorFactura"] != System.DBNull.Value) dValorFactura = decimal.Parse(row["ValorFactura"].ToString());
						if (row["Estado"] != System.DBNull.Value) sEstado = row["Estado"].ToString();						

						string sSQLAct = string.Format("Exec GrabaRegFacDentalis {0}, '{1}', '{2}', '{3}', '{4}', '{5}','{6}','{7}','{8}',{9},'{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}', {18},{19},'{20}','{21}','{22}','{23}','{24}',{25},'{26}'", 
							IdFactDeantalis, sSucursal, sEstablecimiento, Convert.ToDateTime(sFecha).ToString("yyyyMMdd"), sTipoIdentificacion, sRut, sPaciente, sTELEFONO, sDIRECCION, iPresupuesto, sConvenioPresupuesto, sEspecialidadPresupuesto,	sConvenioPac, sInstitucion, sDiagnosticador, sComprobante, sCaja, sBancoTarjeta, iLOTE, iTRANSSACCION, sTURNO,                             sFACTURA, iNumDoc, Convert.ToDateTime(sVcto).ToString("yyyyMMdd"), sDcto_ppto, dValorFactura, sEstado);						
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
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void frmFactDentalis_Load(object sender, System.EventArgs e)
		{
		
		}
	}
}
