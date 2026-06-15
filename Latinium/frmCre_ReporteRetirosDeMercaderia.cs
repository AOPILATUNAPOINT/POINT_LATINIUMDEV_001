using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCre_ReporteRetirosDeMercaderia.
	/// </summary>
	public class frmCre_ReporteRetirosDeMercaderia : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label53;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optFiltro;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCre_ReporteRetirosDeMercaderia()
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
		#region Código generado por el Diseñador de Windows Forms
		/// <summary>
		/// Método necesario para admitir el Diseñador. No se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCre_ReporteRetirosDeMercaderia));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			this.label53 = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnVer = new System.Windows.Forms.Button();
			this.uGridInformacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.optFiltro = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optFiltro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// label53
			// 
			this.label53.AutoSize = true;
			this.label53.BackColor = System.Drawing.Color.Transparent;
			this.label53.Location = new System.Drawing.Point(196, 19);
			this.label53.Name = "label53";
			this.label53.Size = new System.Drawing.Size(34, 16);
			this.label53.TabIndex = 397;
			this.label53.Text = "Hasta";
			this.label53.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.BackColor = System.Drawing.Color.Transparent;
			this.lblBodega.Location = new System.Drawing.Point(19, 19);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(37, 16);
			this.lblBodega.TabIndex = 396;
			this.lblBodega.Text = "Desde";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtHasta
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance1;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(244, 17);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 395;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance2;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(70, 17);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 394;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(370, 15);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 398;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// uGridInformacion
			// 
			this.uGridInformacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridInformacion.Cursor = System.Windows.Forms.Cursors.Default;
			appearance3.BackColor = System.Drawing.Color.White;
			this.uGridInformacion.DisplayLayout.Appearance = appearance3;
			this.uGridInformacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridInformacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridInformacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.uGridInformacion.DisplayLayout.Override.CardAreaAppearance = appearance4;
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInformacion.DisplayLayout.Override.CellAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 8F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInformacion.DisplayLayout.Override.HeaderAppearance = appearance6;
			this.uGridInformacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.uGridInformacion.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInformacion.Location = new System.Drawing.Point(8, 56);
			this.uGridInformacion.Name = "uGridInformacion";
			this.uGridInformacion.Size = new System.Drawing.Size(1320, 568);
			this.uGridInformacion.TabIndex = 668;
			// 
			// optFiltro
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.optFiltro.Appearance = appearance9;
			this.optFiltro.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optFiltro.ItemAppearance = appearance10;
			this.optFiltro.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Cliente";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Articulos";
			this.optFiltro.Items.Add(valueListItem1);
			this.optFiltro.Items.Add(valueListItem2);
			this.optFiltro.ItemSpacingVertical = 10;
			this.optFiltro.Location = new System.Drawing.Point(584, 16);
			this.optFiltro.Name = "optFiltro";
			this.optFiltro.Size = new System.Drawing.Size(152, 26);
			this.optFiltro.TabIndex = 669;
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
			// frmCre_ReporteRetirosDeMercaderia
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1338, 632);
			this.Controls.Add(this.optFiltro);
			this.Controls.Add(this.uGridInformacion);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.label53);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmCre_ReporteRetirosDeMercaderia";
			this.Text = "Retiros De Mercaderia";
			this.Load += new System.EventHandler(this.frmCre_ReporteRetirosDeMercaderia_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optFiltro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			#region Valida Fechas
			if (this.dtDesde.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtDesde.Focus();
				return;
			}

			if (this.dtHasta.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtHasta.Focus();
				return;
			}

			if (Convert.ToDateTime(this.dtDesde.Value) > Convert.ToDateTime(this.dtHasta.Value))
			{
				MessageBox.Show("La Fecha Inicial no puede ser mayor a la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtHasta.Focus();
				return;
			}
			#endregion Valida Fechas

			ParameterFields paramFields = new ParameterFields ();
											
			ParameterField paramFieldCompra = new ParameterField ();
			ParameterDiscreteValue ValIdCompra = new ParameterDiscreteValue ();
			paramFieldCompra.ParameterFieldName = "@Desde";
			ValIdCompra.Value = Convert.ToDateTime(this.dtDesde.Value);
			paramFieldCompra.CurrentValues.Add (ValIdCompra);
			paramFields.Add (paramFieldCompra);
							
			ParameterField paramFieldFecha = new ParameterField ();
			ParameterDiscreteValue ValIdFecha = new ParameterDiscreteValue ();
			paramFieldFecha.ParameterFieldName = "@Hasta";
			ValIdFecha.Value = Convert.ToDateTime(this.dtHasta.Value);
			paramFieldFecha.CurrentValues.Add (ValIdFecha);
			paramFields.Add (paramFieldFecha);
											
			Reporte miRep = new Reporte("Cre_ReporteRetirosDeCredito.rpt", "");
			miRep.MdiParent = this.MdiParent;
			miRep.Titulo("Retiros de Meraderia");
			miRep.crVista.ParameterFieldInfo = paramFields;
			miRep.Show();
		}

		private void frmCre_ReporteRetirosDeMercaderia_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "081209");

			if (!Funcion.Permiso("935", cdsSeteoF))
			{				
				MessageBox.Show("No tiene Acceso RETIROS DE MERCADERIA", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			this.dtDesde.CalendarInfo.MaxDate = DateTime.Today;			
			this.dtHasta.CalendarInfo.MaxDate = DateTime.Today;

			this.dtDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();			
			this.dtHasta.Value = DateTime.Today;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
		

			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	

		}
	}
}
