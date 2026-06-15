using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteAuditoriaLocales.
	/// </summary>
	public class frmReporteAuditoriaLocales : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Label lblBodega;
		private System.Windows.Forms.Label lblfecha;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbAuditor;
		private System.Windows.Forms.Label label1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReporteAuditoriaLocales()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteAuditoriaLocales));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblBodega = new System.Windows.Forms.Label();
			this.lblfecha = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label7 = new System.Windows.Forms.Label();
			this.cmbAuditor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label1 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAuditor)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbBodega
			// 
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(272, 24);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(180, 22);
			this.cmbBodega.TabIndex = 758;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(208, 24);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(43, 16);
			this.lblBodega.TabIndex = 757;
			this.lblBodega.Text = "Bodega";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblfecha
			// 
			this.lblfecha.AutoSize = true;
			this.lblfecha.Location = new System.Drawing.Point(16, 24);
			this.lblfecha.Name = "lblfecha";
			this.lblfecha.Size = new System.Drawing.Size(36, 16);
			this.lblfecha.TabIndex = 755;
			this.lblfecha.Text = "Fecha";
			this.lblfecha.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance1;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(72, 24);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 756;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
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
			// cmbVendedor
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVendedor.Appearance = appearance2;
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 335;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(520, 24);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(320, 21);
			this.cmbVendedor.TabIndex = 759;
			this.cmbVendedor.ValueMember = "idPersonal";
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(456, 24);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(44, 16);
			this.label7.TabIndex = 760;
			this.label7.Text = "Nombre";
			// 
			// cmbAuditor
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbAuditor.Appearance = appearance3;
			this.cmbAuditor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbAuditor.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 335;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbAuditor.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbAuditor.DisplayMember = "Nombre";
			this.cmbAuditor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbAuditor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbAuditor.Location = new System.Drawing.Point(72, 80);
			this.cmbAuditor.Name = "cmbAuditor";
			this.cmbAuditor.Size = new System.Drawing.Size(320, 21);
			this.cmbAuditor.TabIndex = 761;
			this.cmbAuditor.ValueMember = "idPersonal";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 80);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(44, 16);
			this.label1.TabIndex = 762;
			this.label1.Text = "Nombre";
			// 
			// frmReporteAuditoriaLocales
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(848, 430);
			this.Controls.Add(this.cmbAuditor);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbVendedor);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.lblfecha);
			this.Controls.Add(this.dtFecha);
			this.Name = "frmReporteAuditoriaLocales";
			this.Text = "frmReporteAuditoriaLocales";
			this.Load += new System.EventHandler(this.frmReporteAuditoriaLocales_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAuditor)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmReporteAuditoriaLocales_Load(object sender, System.EventArgs e)
		{
			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;
			this.dtFecha.Value = DateTime.Today;
			this.cmbBodega.DataSource = FuncionesProcedimientos.dtGeneral("Select Bodega, Codigo, Nombre From Bodega Where Activo = 1 And Factura = 1 Order By Nombre");
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			int ibodega = (int)this.cmbBodega.Value;
			this.cmbVendedor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "sELECT N.idPersonal,N.ApellidoPaterno+' '+N.ApellidoMaterno+' '+PrimerNombre AS Nombre From Com_AsignacionDeVendedores c Inner Join Bodega b On b.Bodega = c.Bodega iNNER JOIN Nomina N ON N.idPersonal = C.idPersonal Where C.idCargo = 1 AND Periodo = '01/2018' AND(c.Bodega = '"+ ibodega+"')Order By b.Nombre");
}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
		this.cmbAuditor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "sELECT idPersonal, ApellidoPaterno + ' '+ ApellidoMaterno+ ' '+ PrimerNombre+ ' '+ SegundoNombre As Nombre  FROM Nomina WHERE idNomina IN (128,107,256,79)");
		}
	}
}
