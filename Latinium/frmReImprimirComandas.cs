using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using Infragistics.Win.UltraWinGrid;
using CrystalDecisions.Shared;
using C1.Data;
using System.Drawing.Printing;
using System.IO;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReImprimirComandas.
	/// </summary>
	public class frmReImprimirComandas : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private C1.Data.C1DataSet cdsSeteoF;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridDetalle;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReImprimirComandas()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}
		int idCliente = 0;
		int iBodega = 0;
		public frmReImprimirComandas(int IdCliente,int IBodega)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
		
			idCliente = IdCliente;
			iBodega = IBodega;

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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idComanda");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroComanda");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idHabitacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Encuesta");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idComanda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroComanda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idHabitacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Encuesta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnImpresion", 0);
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReImprimirComandas));
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.uGridDetalle = new Infragistics.Win.UltraWinGrid.UltraGrid();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraDataSource3
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(System.Decimal);
			ultraDataColumn5.DataType = typeof(System.Decimal);
			ultraDataColumn12.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn12});
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
			// uGridDetalle
			// 
			this.uGridDetalle.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDetalle.DataSource = this.ultraDataSource3;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridDetalle.DisplayLayout.Appearance = appearance1;
			this.uGridDetalle.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 52;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.Caption = "Numero Comanda";
			ultraGridColumn2.Header.VisiblePosition = 3;
			ultraGridColumn2.Width = 117;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 5;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 75;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 6;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 70;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 7;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 86;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 8;
			ultraGridColumn6.Width = 109;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 4;
			ultraGridColumn8.Width = 345;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 9;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 89;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.Caption = "Habitacion";
			ultraGridColumn10.Header.VisiblePosition = 10;
			ultraGridColumn10.Width = 104;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 11;
			ultraGridColumn11.Width = 112;
			ultraGridColumn12.Header.VisiblePosition = 12;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance2.Image = ((object)(resources.GetObject("appearance2.Image")));
			appearance2.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance2.ImageVAlign = Infragistics.Win.VAlign.Middle;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn13.CellAppearance = appearance2;
			appearance3.Image = ((object)(resources.GetObject("appearance3.Image")));
			appearance3.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance3.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn13.CellButtonAppearance = appearance3;
			ultraGridColumn13.Header.Caption = "Comanda";
			ultraGridColumn13.Header.VisiblePosition = 2;
			ultraGridColumn13.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn13.Width = 48;
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
																										 ultraGridColumn13});
			ultraGridBand1.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.uGridDetalle.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalle.DisplayLayout.Override.ActiveRowAppearance = appearance4;
			this.uGridDetalle.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridDetalle.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridDetalle.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance5.BackColor = System.Drawing.Color.Transparent;
			this.uGridDetalle.DisplayLayout.Override.CardAreaAppearance = appearance5;
			appearance6.ForeColor = System.Drawing.Color.Black;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalle.DisplayLayout.Override.CellAppearance = appearance6;
			this.uGridDetalle.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance7.FontData.BoldAsString = "True";
			appearance7.FontData.Name = "Arial";
			appearance7.FontData.SizeInPoints = 9F;
			appearance7.ForeColor = System.Drawing.Color.White;
			appearance7.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDetalle.DisplayLayout.Override.HeaderAppearance = appearance7;
			appearance8.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance8.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowAlternateAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.uGridDetalle.Dock = System.Windows.Forms.DockStyle.Fill;
			this.uGridDetalle.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDetalle.Location = new System.Drawing.Point(0, 0);
			this.uGridDetalle.Name = "uGridDetalle";
			this.uGridDetalle.Size = new System.Drawing.Size(856, 342);
			this.uGridDetalle.TabIndex = 299;
			this.uGridDetalle.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDetalle_ClickCellButton);
			// 
			// frmReImprimirComandas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(856, 342);
			this.Controls.Add(this.uGridDetalle);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmReImprimirComandas";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "ReImprimir Comandas";
			this.Load += new System.EventHandler(this.frmReImprimirComandas_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmReImprimirComandas_Load(object sender, System.EventArgs e)
		{
			string sSQL = string.Format(string.Format("Exec ConsultaComandaReimprimir {0}", idCliente));
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridDetalle);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void uGridDetalle_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "btnImpresion")
			{
				if (e.Cell.Row.Cells["idComanda"].Value == null || e.Cell.Row.Cells["idComanda"].Value == DBNull.Value)
				{
					MessageBox.Show("No tiene comandas para imprimir.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}

				int idComanda = (int)e.Cell.Row.Cells["idComanda"].Value;
				int Encuesta = (int)e.Cell.Row.Cells["Encuesta"].Value;
				if (idComanda == 0)
				{
					MessageBox.Show("No tiene comandas para imprimir.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec ActualizaEstadoComandaEncuesta {0}", idComanda));

				string sImpresora = "";
				sImpresora = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" exec ConsultarBodegaProservicios {0}, 1 ",iBodega));		
			              
				string sImpresoraCocina = "";
				sImpresoraCocina = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" exec ConsultarBodegaProservicios {0}, 2 ",iBodega));		

				ImprimeComandaBtnOrdenCaja(idComanda, sImpresora);
				ImprimeComandaBtnOrdenCliente(idComanda, sImpresoraCocina);

			}
		}
		private void ImprimeComandaBtnOrdenCaja(int iComanda, string sImpresora)
		{
			#region Impresion Comanda Caja
			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramFieldIdDetCompra = new ParameterField ();
			ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
			paramFieldIdDetCompra.ParameterFieldName = "@idComanda";
			discreteValIdDetCompra.Value = iComanda; 
			paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
			paramFields.Add (paramFieldIdDetCompra);
				
			PrintDocument pd = new PrintDocument();
			pd.PrinterSettings.PrinterName = sImpresora;
			Reporte Reporte = new Reporte("ImpresionComandabtnOrdenEncuesta.rpt", "");
			Reporte.oRpt.PrintOptions.PrinterName = sImpresora;			  
			Reporte.MdiParent = this.MdiParent;
			Reporte.Titulo("Impresion Comanda Caja");		 
			Reporte.crVista.ParameterFieldInfo = paramFields;
			Reporte.ShowDialog();
			Reporte.Show();
			#endregion Impresion Comanda Caja
		}
		private void ImprimeComandaBtnOrdenCliente(int iComanda, string sImpresora)
		{
			#region Impresion Comanda Cliente
			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramFieldIdDetCompra = new ParameterField ();
			ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
			paramFieldIdDetCompra.ParameterFieldName = "@idComanda";
			discreteValIdDetCompra.Value = iComanda; 
			paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
			paramFields.Add (paramFieldIdDetCompra);
				
			PrintDocument pd = new PrintDocument();
			pd.PrinterSettings.PrinterName = sImpresora;
			Reporte Reporte = new Reporte("ImpresionComandabtnOrdenEncuesta.rpt", "");
			Reporte.oRpt.PrintOptions.PrinterName = sImpresora;			  
			Reporte.MdiParent = this.MdiParent;
			Reporte.Titulo("Impresion Comanda Cocina");		 
			Reporte.crVista.ParameterFieldInfo = paramFields;
			Reporte.ShowDialog();
			Reporte.Show();
			#endregion Impresion Comanda Cliente
		}
	}
}
