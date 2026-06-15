using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.IO;
using System.Drawing;
using System.Data.SqlClient;
using System.Data;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmVistaPromociones.
	/// </summary>
	public class frmVistaPromociones : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label lblVendedor;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFin;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtInicio;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.PictureBox pbImagen;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPromocion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.RichTextBox rtbDescripcion;
		private System.Windows.Forms.PictureBox pbCompetencia;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmVistaPromociones()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDescripcionPromociones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDescripcionPromociones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmVistaPromociones));
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.lblVendedor = new System.Windows.Forms.Label();
			this.cmbPromocion = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label6 = new System.Windows.Forms.Label();
			this.dtFin = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtInicio = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.pbImagen = new System.Windows.Forms.PictureBox();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.rtbDescripcion = new System.Windows.Forms.RichTextBox();
			this.pbCompetencia = new System.Windows.Forms.PictureBox();
			((System.ComponentModel.ISupportInitialize)(this.cmbPromocion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFin)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtInicio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// lblVendedor
			// 
			this.lblVendedor.AutoSize = true;
			this.lblVendedor.Location = new System.Drawing.Point(8, 8);
			this.lblVendedor.Name = "lblVendedor";
			this.lblVendedor.Size = new System.Drawing.Size(57, 16);
			this.lblVendedor.TabIndex = 153;
			this.lblVendedor.Text = "Promocion";
			// 
			// cmbPromocion
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPromocion.Appearance = appearance1;
			this.cmbPromocion.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPromocion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPromocion.DataSource = this.ultraDataSource1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 520;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbPromocion.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbPromocion.DisplayMember = "Nombre";
			this.cmbPromocion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPromocion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPromocion.Location = new System.Drawing.Point(80, 8);
			this.cmbPromocion.Name = "cmbPromocion";
			this.cmbPromocion.Size = new System.Drawing.Size(520, 21);
			this.cmbPromocion.TabIndex = 152;
			this.cmbPromocion.ValueMember = "idDescripcionPromociones";
			this.cmbPromocion.ValueChanged += new System.EventHandler(this.cmbPromocion_ValueChanged);
			this.cmbPromocion.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbPromocion_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Enabled = false;
			this.label6.Location = new System.Drawing.Point(8, 40);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(35, 16);
			this.label6.TabIndex = 229;
			this.label6.Text = "Desde";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFin
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFin.Appearance = appearance2;
			this.dtFin.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFin.DateButtons.Add(dateButton1);
			this.dtFin.Enabled = false;
			this.dtFin.Format = "dd/MM/yyyy HH:mm";
			this.dtFin.Location = new System.Drawing.Point(296, 40);
			this.dtFin.Name = "dtFin";
			this.dtFin.NonAutoSizeHeight = 23;
			this.dtFin.Size = new System.Drawing.Size(144, 21);
			this.dtFin.SpinButtonsVisible = true;
			this.dtFin.TabIndex = 228;
			this.dtFin.Value = ((object)(resources.GetObject("dtFin.Value")));
			// 
			// dtInicio
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtInicio.Appearance = appearance3;
			this.dtInicio.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtInicio.DateButtons.Add(dateButton2);
			this.dtInicio.Enabled = false;
			this.dtInicio.Format = "dd/MM/yyyy HH:mm";
			this.dtInicio.Location = new System.Drawing.Point(80, 40);
			this.dtInicio.Name = "dtInicio";
			this.dtInicio.NonAutoSizeHeight = 23;
			this.dtInicio.Size = new System.Drawing.Size(144, 21);
			this.dtInicio.SpinButtonsVisible = true;
			this.dtInicio.TabIndex = 227;
			this.dtInicio.Value = ((object)(resources.GetObject("dtInicio.Value")));
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Enabled = false;
			this.label1.Location = new System.Drawing.Point(248, 40);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(33, 16);
			this.label1.TabIndex = 232;
			this.label1.Text = "Hasta";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// pbImagen
			// 
			this.pbImagen.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.pbImagen.Enabled = false;
			this.pbImagen.Location = new System.Drawing.Point(605, 72);
			this.pbImagen.Name = "pbImagen";
			this.pbImagen.Size = new System.Drawing.Size(400, 520);
			this.pbImagen.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
			this.pbImagen.TabIndex = 233;
			this.pbImagen.TabStop = false;
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
			// rtbDescripcion
			// 
			this.rtbDescripcion.Font = new System.Drawing.Font("Tahoma", 8F);
			this.rtbDescripcion.Location = new System.Drawing.Point(5, 72);
			this.rtbDescripcion.Name = "rtbDescripcion";
			this.rtbDescripcion.ReadOnly = true;
			this.rtbDescripcion.Size = new System.Drawing.Size(592, 520);
			this.rtbDescripcion.TabIndex = 234;
			this.rtbDescripcion.Text = "";
			// 
			// pbCompetencia
			// 
			this.pbCompetencia.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.pbCompetencia.Enabled = false;
			this.pbCompetencia.Location = new System.Drawing.Point(1016, 72);
			this.pbCompetencia.Name = "pbCompetencia";
			this.pbCompetencia.Size = new System.Drawing.Size(400, 520);
			this.pbCompetencia.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
			this.pbCompetencia.TabIndex = 235;
			this.pbCompetencia.TabStop = false;
			// 
			// frmVistaPromociones
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.AutoScroll = true;
			this.ClientSize = new System.Drawing.Size(1426, 600);
			this.Controls.Add(this.pbCompetencia);
			this.Controls.Add(this.rtbDescripcion);
			this.Controls.Add(this.pbImagen);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.dtFin);
			this.Controls.Add(this.dtInicio);
			this.Controls.Add(this.lblVendedor);
			this.Controls.Add(this.cmbPromocion);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmVistaPromociones";
			this.ShowInTaskbar = false;
			this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show;
			this.Text = "Promociones Activas";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmVistaPromociones_KeyDown);
			this.Load += new System.EventHandler(this.frmVistaPromociones_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbPromocion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFin)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtInicio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void frmVistaPromociones_Load(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Select idDescripcionPromociones, Nombre From DescripcionPromociones Where Estado = 1 Order By idDescripcionPromociones Desc");
			this.cmbPromocion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
						
			this.cmbPromocion.Value = Funcion.iEscalarSQL(cdsSeteoF, "Declare @id Int = 0 If Exists(Select idDescripcionPromociones From DescripcionPromociones Where Estado = 1)Select Top 1 @id = idDescripcionPromociones From DescripcionPromociones Where Estado = 1 Order By idDescripcionPromociones Desc Select @id");

			sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);
			int iBodegaPredefinida = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			if (iBodegaPredefinida > 0)
			{
				sSQL = string.Format("Exec RetornaImagenCompetencia {0}, '{1}'", iBodegaPredefinida, DateTime.Today.ToString("yyyyMMdd"));
				string sArchivo = Funcion.sEscalarSQL(cdsSeteoF, sSQL);

				if (File.Exists(sArchivo))
					pbCompetencia.Image = Image.FromFile(sArchivo);			
			}
		}

		private void cmbPromocion_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void Consultar(int idDescripcion)
		{
			string sSQL = string.Format("Exec BuscaDescripcionPromociones '', {0}", idDescripcion);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.dtInicio.Value = dr.GetDateTime(2);
				this.dtFin.Value = dr.GetDateTime(3);				
				this.rtbDescripcion.Text = dr.GetString(4);
				byte[] arrImg = (byte[])dr.GetValue(5);
					
				MemoryStream ms = new MemoryStream(arrImg);
				this.pbImagen.Image = System.Drawing.Image.FromStream(ms);
				ms.Close();
			}
			dr.Close();
		}

		private void cmbPromocion_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbPromocion.ActiveRow != null) this.Consultar((int)this.cmbPromocion.Value);
		}

		private void frmVistaPromociones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}
	}
}

