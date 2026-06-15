using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Text;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmLoteBancoDelPichincha.
	/// </summary>
	public class frmLoteBancoDelPichincha : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Button btnExportarTXT;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBusqueda;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtLote;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridLotes;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnBuscar;
		private System.Windows.Forms.Label label1;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Label lblContador;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmLoteBancoDelPichincha()
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsientoLoteSolidario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CapturaLote");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroLote");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsientoLoteSolidario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CapturaLote");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroLote");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmLoteBancoDelPichincha));
			this.btnExportarTXT = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label4 = new System.Windows.Forms.Label();
			this.txtBusqueda = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtLote = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridLotes = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label1 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblContador = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLote)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLotes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			this.SuspendLayout();
			// 
			// btnExportarTXT
			// 
			this.btnExportarTXT.CausesValidation = false;
			this.btnExportarTXT.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportarTXT.Location = new System.Drawing.Point(752, 8);
			this.btnExportarTXT.Name = "btnExportarTXT";
			this.btnExportarTXT.Size = new System.Drawing.Size(96, 23);
			this.btnExportarTXT.TabIndex = 286;
			this.btnExportarTXT.Text = "Exportar a TXT";
			this.btnExportarTXT.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportarTXT.Click += new System.EventHandler(this.btnExportarTXT_Click);
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
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(184, 8);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(38, 16);
			this.label4.TabIndex = 629;
			this.label4.Text = "Cliente";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtBusqueda
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBusqueda.Appearance = appearance1;
			this.txtBusqueda.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBusqueda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBusqueda.Location = new System.Drawing.Point(240, 8);
			this.txtBusqueda.Name = "txtBusqueda";
			this.txtBusqueda.Size = new System.Drawing.Size(176, 22);
			this.txtBusqueda.TabIndex = 628;
			// 
			// txtLote
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtLote.Appearance = appearance2;
			this.txtLote.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtLote.FormatString = "";
			this.txtLote.Location = new System.Drawing.Point(64, 8);
			this.txtLote.Name = "txtLote";
			this.txtLote.PromptChar = ' ';
			this.txtLote.Size = new System.Drawing.Size(96, 22);
			this.txtLote.TabIndex = 627;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(8, 8);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(40, 16);
			this.label3.TabIndex = 626;
			this.label3.Text = "N. Lote";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(bool);
			ultraDataColumn3.DefaultValue = false;
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(System.DateTime);
			ultraDataColumn6.DefaultValue = new System.DateTime(((long)(0)));
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8});
			// 
			// uGridLotes
			// 
			this.uGridLotes.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridLotes.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridLotes.DataSource = this.ultraDataSource1;
			appearance3.BackColor = System.Drawing.Color.White;
			this.uGridLotes.DisplayLayout.Appearance = appearance3;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 368;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 44;
			ultraGridColumn3.Header.Caption = "...";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 25;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance4;
			ultraGridColumn4.Header.Caption = "Lote";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 100;
			ultraGridColumn5.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "N. Operación";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 120;
			ultraGridColumn6.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 75;
			ultraGridColumn7.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.Caption = "Cédula";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 120;
			ultraGridColumn8.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 379;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridLotes.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLotes.DisplayLayout.GroupByBox.Appearance = appearance5;
			appearance6.ForeColor = System.Drawing.Color.Black;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLotes.DisplayLayout.GroupByBox.BandLabelAppearance = appearance6;
			this.uGridLotes.DisplayLayout.GroupByBox.Hidden = true;
			this.uGridLotes.DisplayLayout.GroupByBox.Prompt = " ";
			appearance7.BackColor = System.Drawing.SystemColors.Control;
			appearance7.BackColor2 = System.Drawing.SystemColors.Control;
			appearance7.BackColorDisabled = System.Drawing.SystemColors.Control;
			appearance7.BackColorDisabled2 = System.Drawing.SystemColors.Control;
			appearance7.ForeColor = System.Drawing.Color.Black;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLotes.DisplayLayout.GroupByBox.PromptAppearance = appearance7;
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLotes.DisplayLayout.Override.ActiveRowAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.Transparent;
			this.uGridLotes.DisplayLayout.Override.CardAreaAppearance = appearance9;
			this.uGridLotes.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.AllRowsInBand;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance10.FontData.BoldAsString = "True";
			appearance10.FontData.Name = "Arial";
			appearance10.FontData.SizeInPoints = 10F;
			appearance10.ForeColor = System.Drawing.Color.White;
			appearance10.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridLotes.DisplayLayout.Override.HeaderAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLotes.DisplayLayout.Override.RowSelectorAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLotes.DisplayLayout.Override.SelectedRowAppearance = appearance12;
			this.uGridLotes.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridLotes.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridLotes.Location = new System.Drawing.Point(8, 72);
			this.uGridLotes.Name = "uGridLotes";
			this.uGridLotes.Size = new System.Drawing.Size(840, 304);
			this.uGridLotes.TabIndex = 630;
			// 
			// btnBuscar
			// 
			this.btnBuscar.CausesValidation = false;
			this.btnBuscar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnBuscar.Location = new System.Drawing.Point(616, 8);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.Size = new System.Drawing.Size(64, 23);
			this.btnBuscar.TabIndex = 631;
			this.btnBuscar.Text = "Buscar";
			this.btnBuscar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 56);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(856, 8);
			this.groupBox1.TabIndex = 632;
			this.groupBox1.TabStop = false;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(432, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(34, 16);
			this.label1.TabIndex = 634;
			this.label1.Text = "Fecha";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance13;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton1.Caption = "Today";
			dateButton1.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(480, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(104, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 633;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.BackColor = System.Drawing.Color.Transparent;
			this.lblContador.Location = new System.Drawing.Point(8, 32);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 16);
			this.lblContador.TabIndex = 635;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// frmLoteBancoDelPichincha
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(856, 382);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.txtBusqueda);
			this.Controls.Add(this.txtLote);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.btnExportarTXT);
			this.Controls.Add(this.uGridLotes);
			this.Name = "frmLoteBancoDelPichincha";
			this.Text = "Lote Banco Del Pichincha";
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLote)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLotes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void ExportaClientes ()
		{			
			/*
			INDACCIONIDENT CodEmpresa CodAgencia CodDocOficial CodNumDocOficial NomApellido1 6
			NomApellido2 NomPersona1 NomPersona2 CodNacionalidad FecExpiraDoc FecIniRes 12
			PaisNacimnto CodProvincia CodMunicipio CodSexo CodEstadoCivil CodNivelEstudios FechaNacimiento 19
			IndVivienda NumCargas CodSitLaboral CodCNAE NOMDENOMSOCIALA FecIngreso CODSECTORBANCOE 26
			IngresosMensuales EgresosMensuales TotalActivos TotalPasivos CodTipoDirec CodPaisD CodProvincia 33
			CodMunicipio NomVia MasDetalle NOMAMPLDOMICILIO PrefijoIntTF PrefijoTF NumTelefono NumCelular Email 42
			CodReferencia1 NumReferencia1 NomRefPersonal1 NUMTFNOREF1 CODPROVINCIA1 CodMunicipio1 CodReferencia2 NumReferencia2 50
			NomRefPersonal2 NUMTFNOREF2 CODPROVINCIA2 CodMunicipio2 CodCanton CodParroquia TipoEmpresa ProRelComercial 58
			*/ 

			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
			{	
				DateTime dtFecha = (DateTime)this.dtFecha.Value;

				SqlCommand oCmdActualiza = oConexion.CreateCommand();
				oCmdActualiza.Connection = oConexion;
				oCmdActualiza.CommandTimeout = 0;
				string sSQL = string.Format("Exec DatBanco_Pichincha 1, '{0}'", dtFecha.ToString("yyyyMMdd"));
				oCmdActualiza.CommandText = sSQL;
				SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
				DataTable oDTDetalle = new DataTable();
				oDADetalle.Fill(oDTDetalle);						
						
				using(StreamWriter file = new StreamWriter("CLIPNT.txt", true))
				{
					int i = 0;
					foreach(DataRow dr in oDTDetalle.Rows)
					{					
						i++;
		
						if (dr["CodNumDocOficial"].ToString().Length != 14)
						{
							MessageBox.Show(string.Format("La Longitud del campo Cedula debe ser de 14 - {0}", dr["CodNumDocOficial"].ToString().Length));
							return;
						}

						if (dr["NumTelefono"].ToString().Length != 10)
						{
							MessageBox.Show(string.Format("La Longitud del campo Telefono domicilio  debe ser de 10 - {0}", dr["NumTelefono"].ToString().Length));
							return;
						}

						if (dr["NumCelular"].ToString().Length != 10)
						{
							MessageBox.Show(string.Format("La Longitud del campo Celular debe ser de 10 - {0}", dr["NumCelular"].ToString().Length));
							return;
						}

						if (dr["NUMTFNOREF1"].ToString().Length != 10)
						{
							MessageBox.Show(string.Format("La Longitud del campo Telefono referencia 1 debe ser de 10 - {0}", dr["NUMTFNOREF1"].ToString().Length));
							return;
						}

						if (dr["NUMTFNOREF2"].ToString().Length != 10)
						{
							MessageBox.Show(string.Format("La Longitud del campo Telefono referencia 2 debe ser de 10 - {0}", dr["NUMTFNOREF2"].ToString().Length));
							return;
						}

						if (dr["NOMDENOMSOCIALA"].ToString() == "")
						{
							MessageBox.Show(string.Format("El campo Empresa Trabajo No puede estar vacio ", dr["NUMTFNOREF2"].ToString().Length));
							return;
						}

						string iLinea = dr["INDACCIONIDENT"].ToString() + dr["CodEmpresa"].ToString() + dr["CodAgencia"].ToString() + dr["CodDocOficial"].ToString() + dr["CodNumDocOficial"].ToString() + dr["NomApellido1"].ToString() 
							+ dr["NomApellido2"].ToString() + dr["NomPersona1"].ToString() + dr["NomPersona2"].ToString() + dr["CodNacionalidad"].ToString() + dr["FecExpiraDoc"].ToString() + dr["FecIniRes"].ToString() 
							+ dr["PaisNacimnto"].ToString() + dr["CodProvincia"].ToString() + dr["CodMunicipio"].ToString() + dr["CodSexo"].ToString() + dr["CodEstadoCivil"].ToString() + dr["CodNivelEstudios"].ToString() + dr["FechaNacimiento"].ToString() 
							+ dr["IndVivienda"].ToString() + dr["NumCargas"].ToString() + dr["CodSitLaboral"].ToString() + dr["CodCNAE"].ToString() + dr["NOMDENOMSOCIALA"].ToString() + dr["FecIngreso"].ToString() + dr["CODSECTORBANCOE"].ToString()
							+ dr["IngresosMensuales"].ToString() + dr["EgresosMensuales"].ToString() + dr["TotalActivos"].ToString() + dr["TotalPasivos"].ToString() + dr["CodTipoDirec"].ToString() + dr["CodPaisD"].ToString() + dr["CodProvinciaDom"].ToString()
							+ dr["CodMunicipioDom"].ToString() + dr["NomVia"].ToString() + dr["MasDetalle"].ToString() + dr["NOMAMPLDOMICILIO"].ToString() + dr["PrefijoIntTF"].ToString() + dr["PrefijoTF"].ToString() + dr["NumTelefono"].ToString() + dr["NumCelular"].ToString() + dr["Email"].ToString()
							+ dr["CodReferencia1"].ToString() + dr["NumReferencia1"].ToString() + dr["NomRefPersonal1"].ToString() + dr["NUMTFNOREF1"].ToString() + dr["CODPROVINCIA1"].ToString() + dr["CodMunicipio1"].ToString() + dr["CodReferencia2"].ToString() + dr["NumReferencia2"].ToString()
							+ dr["NomRefPersonal2"].ToString() + dr["NUMTFNOREF2"].ToString() + dr["CODPROVINCIA2"].ToString() + dr["CodMunicipio2"].ToString() + dr["CodCanton"].ToString() + dr["CodParroquia"].ToString() + dr["TipoEmpresa"].ToString() + dr["ProRelComercial"].ToString();
									
						if (iLinea.Length != 622)
						{
							MessageBox.Show(string.Format("El registro N. {0} del Cliente '{2}' no cumple con la Longitud requerida (622) - {1}", i, iLinea.Length, dr["CodNumDocOficial"].ToString()));
							return;
						}

						file.WriteLine(iLinea);
					}					
				}
			}				
		}

		int iContOperaciones = 0;

		private void ExportaGAF()
		{
			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
			{			
				DateTime dtFecha = (DateTime)this.dtFecha.Value;

				SqlCommand oCmdActualiza = oConexion.CreateCommand();
				oCmdActualiza.Connection = oConexion;
				oCmdActualiza.CommandTimeout = 0;
				string sSQL = string.Format("Exec ListaClientesGAFBDP 4, '{0}'", dtFecha.ToString("yyyyMMdd"));
				oCmdActualiza.CommandText = sSQL;
				SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
				DataTable oDTDetalle = new DataTable();
				oDADetalle.Fill(oDTDetalle);						
				
				iContOperaciones = 0;
				//Fecha NumeroLote Peticion Producto Referencia AgenciaIngreso Usuario FechaIngreso HoraIngreso SistemaEntrada
				//Datos_Cliente Valor_Prestamo Fecha_Efectiva Fecha_Vencimiento Plazo PeriodoPlazo PeriodoPago TasaInteres FormaVencDividendo CriterioVencDividendo TipoCredito DatAgencia Datos_Imp_Com GaranteCodeudor DatTblAmor FechaNego
				//TotalCuotas TotalCapital TotalInteres Origen
				
				using(StreamWriter file = new StreamWriter("GAFPNT.txt", true))
				{
					foreach(DataRow dr in oDTDetalle.Rows)
					{	
						if ((Convert.ToDecimal(dr["TotalCuotas"]) - (Convert.ToDecimal(dr["TotalCapital"]) + Convert.ToDecimal(dr["TotalInteres"]))) != 0.00m)
						{
							MessageBox.Show(string.Format("En el Cliente Con Cedula '{0}' Existe Diferencia de {1} Origen '{2}'", 
								dr["Datos_Cliente"].ToString(), Convert.ToDecimal(dr["TotalCuotas"]) - (Convert.ToDecimal(dr["TotalCapital"]) + Convert.ToDecimal(dr["TotalInteres"])), dr["Origen"].ToString()));						
						}

						string sGarante = "";
											
						if (dr["GaranteCodeudor"].ToString().Length > 0) sGarante = dr["GaranteCodeudor"].ToString() + ";";

						string iLinea = dr["Fecha"].ToString() + "\t" + dr["NumeroLote"].ToString() + "\t" + dr["Peticion"].ToString() + "\t" + dr["Producto"].ToString() + "\t" + dr["Referencia"].ToString() + "\t" + dr["AgenciaIngreso"].ToString() + "\t" + dr["Usuario"].ToString() + "\t" + dr["FechaIngreso"].ToString() + "\t" + dr["HoraIngreso"].ToString() + "\t" + dr["Sistemaentrada"].ToString() + "\t" + dr["Datos_Cliente"].ToString() + ";" + dr["Valor_Prestamo"].ToString() + ";" + dr["Fecha_Efectiva"].ToString() + ";" + dr["Fecha_Vencimiento"].ToString() + ";" + dr["Plazo"].ToString() + ";" + dr["PeriodoPlazo"].ToString() + ";" + dr["PeriodoPago"].ToString() + ";" + dr["TasaInteres"].ToString() + ";" + dr["FormaVencDividendo"].ToString() + ";" + dr["CriterioVencDividendo"].ToString() + ";" + dr["TipoCredito"].ToString() + ";" + dr["DatAgencia"].ToString() + ";" + dr["Datos_Imp_Com"].ToString() + ";" + sGarante + dr["DatTblAmor"].ToString() + ";" + dr["FechaNego"].ToString() + ";" + dr["ValDesembolso"].ToString() + "|";
												
						file.WriteLine(iLinea);

						iContOperaciones++;
					}					
				}
			}
		}

		private void btnExportarTXT_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;

			//ExportaClientes();

			ExportaGAF();

			Cursor = Cursors.Default;

			MessageBox.Show(string.Format("Archivo de Operaciones Generado con {0} registros", iContOperaciones));
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			DateTime dtFecha = (DateTime)this.dtFecha.Value;

			bool bFechas = true;

			DateTime dtDesde = DateTime.Today.AddDays(-365) ;
			DateTime dtHasta = DateTime.Today;
			
			bFechas = true;

			dtHasta = (DateTime)this.dtFecha.Value;
		
			string sSQL = string.Format("Exec [ConsultaLotePichincha] '{0}', {1}, '{2}', 4", dtHasta.ToString("yyyyMMdd"), (int)this.txtLote.Value, this.txtBusqueda.Text.ToString());
			this.uGridLotes.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			this.lblContador.Text = this.uGridLotes.Rows.Count + " Registros Encontrados";
		}
	}
}

