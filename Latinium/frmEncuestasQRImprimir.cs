using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using CrystalDecisions.Shared;
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;
using System.IO;
using System.Diagnostics;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmEncuestasQRImprimir.
	/// </summary>
	public class frmEncuestasQRImprimir : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label lblLocal;
		private System.Windows.Forms.Label lblHabitacion;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnImprimir;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugHabitaciones;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmEncuestasQRImprimir()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmEncuestasQRImprimir));
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idHabitacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSucursal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroHabitacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tematica");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PReserva");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idHabitacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SelQR");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroHabitacion", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idHabitacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SelQR");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroHabitacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			this.lblLocal = new System.Windows.Forms.Label();
			this.lblHabitacion = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ugHabitaciones = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugHabitaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			this.SuspendLayout();
			// 
			// lblLocal
			// 
			this.lblLocal.AutoSize = true;
			this.lblLocal.Location = new System.Drawing.Point(24, 24);
			this.lblLocal.Name = "lblLocal";
			this.lblLocal.Size = new System.Drawing.Size(31, 16);
			this.lblLocal.TabIndex = 708;
			this.lblLocal.Text = "Local";
			this.lblLocal.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblHabitacion
			// 
			this.lblHabitacion.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.lblHabitacion.Font = new System.Drawing.Font("Tahoma", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblHabitacion.ForeColor = System.Drawing.Color.Firebrick;
			this.lblHabitacion.Location = new System.Drawing.Point(16, 16);
			this.lblHabitacion.Name = "lblHabitacion";
			this.lblHabitacion.Size = new System.Drawing.Size(448, 40);
			this.lblHabitacion.TabIndex = 706;
			this.lblHabitacion.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
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
			// btnImprimir
			// 
			this.btnImprimir.BackColor = System.Drawing.SystemColors.Control;
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnImprimir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(376, 24);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 24);
			this.btnImprimir.TabIndex = 709;
			this.btnImprimir.Text = "Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(System.Decimal);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			// 
			// cmbBodega
			// 
			appearance1.FontData.Name = "Tahoma";
			this.cmbBodega.Appearance = appearance1;
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 180;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(64, 24);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(250, 22);
			this.cmbBodega.TabIndex = 710;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			// 
			// ugHabitaciones
			// 
			this.ugHabitaciones.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugHabitaciones.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugHabitaciones.DataSource = this.ultraDataSource2;
			appearance2.BackColor = System.Drawing.Color.White;
			this.ugHabitaciones.DisplayLayout.Appearance = appearance2;
			this.ugHabitaciones.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 88;
			ultraGridColumn5.Header.Caption = "...";
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Width = 26;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Width = 334;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn7.Header.Caption = "#";
			ultraGridColumn7.Header.VisiblePosition = 3;
			ultraGridColumn7.Width = 67;
			ultraGridColumn8.Header.VisiblePosition = 4;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 79;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			ultraGridBand2.Header.Caption = "Detalle";
			ultraGridBand2.HeaderVisible = true;
			ultraGridBand2.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugHabitaciones.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ugHabitaciones.DisplayLayout.GroupByBox.Hidden = true;
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugHabitaciones.DisplayLayout.Override.ActiveRowAppearance = appearance3;
			this.ugHabitaciones.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ugHabitaciones.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugHabitaciones.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.ugHabitaciones.DisplayLayout.Override.CardAreaAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 10F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugHabitaciones.DisplayLayout.Override.HeaderAppearance = appearance5;
			this.ugHabitaciones.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance6.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance6.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugHabitaciones.DisplayLayout.Override.RowAlternateAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugHabitaciones.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugHabitaciones.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.ugHabitaciones.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ugHabitaciones.FlatMode = true;
			this.ugHabitaciones.Font = new System.Drawing.Font("Tahoma", 7.75F);
			this.ugHabitaciones.Location = new System.Drawing.Point(16, 72);
			this.ugHabitaciones.Name = "ugHabitaciones";
			this.ugHabitaciones.Size = new System.Drawing.Size(448, 264);
			this.ugHabitaciones.TabIndex = 763;
			this.ugHabitaciones.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugHabitaciones_AfterCellUpdate);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(bool);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11});
			// 
			// frmEncuestasQRImprimir
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(480, 358);
			this.Controls.Add(this.ugHabitaciones);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.lblLocal);
			this.Controls.Add(this.lblHabitacion);
			this.Name = "frmEncuestasQRImprimir";
			this.Text = ".";
			this.Load += new System.EventHandler(this.frmEncuestasQRImprimir_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugHabitaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		#region variables
		int idBodg = 0;
		decimal PorcentajeIva = 0.00m;
		int idRes = 0;
		int idnHbtcn = 0;
		public int idEncuestaConfiguracionQRGeneracion = 1;
		#endregion variables

		private void frmEncuestasQRImprimir_Load(object sender, System.EventArgs e)
		{
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaBodegas {0}, 1", MenuLatinium.IdUsuario));
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			idBodg = (int)this.cmbBodega.Value;
			string ugSQL = string.Format("Exec EncuestasSelQR {0}",idBodg);
			ugHabitaciones.DataSource = Funcion.dvProcedimiento(cdsSeteoF, ugSQL);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void ugHabitaciones_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			string sColumna = e.Cell.Column.Header.Caption;

			if (e.Cell.Column.Header.Caption.ToString() == sColumna)
			{
				EncuestasSelQRActualizar((int)e.Cell.Row.Cells["idHabitacion"].Value,(bool)e.Cell.Row.Cells["SelQR"].Value);

			}
		}

		private void EncuestasSelQRActualizar(int idHabitacion, bool SelQR)
		{
			string squery = "";
			squery = string.Format("Exec EncuestasSelQRActualizar {0}, {1}", idHabitacion,SelQR);
			Funcion.EjecutaSQL(cdsSeteoF, squery, false);
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{

			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione una bodega")) return;
			idBodg = (int)this.cmbBodega.Value;

			if (this.ugHabitaciones.Rows.Count == 0)
			{
				MessageBox.Show("No existen registros para Imprimir", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("EXEC EncuestasSelQRValXBodega {0}", idBodg)) == 1)
			{
				MessageBox.Show("Seleccione la habitación para poder imprimir el código QR", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			int idHabitacion = Funcion.iEscalarSQL(cdsSeteoF, string.Format("EXEC EncuestasSelQRActivo {0}", idBodg));
			string apiService = Funcion.sEscalarSQL(cdsSeteoF, string.Format("EXEC EncuestasConfiguracionModel '{0}', {1}", "Descripcion", idEncuestaConfiguracionQRGeneracion));

			// Generar la URL del servicio
			string urlServicio = string.Format("{0}encuestas/generarQr?bodega={1}&idHabitacion={2}",
				apiService,
				(int)this.cmbBodega.Value,
				idHabitacion
				);

			if (urlServicio.Length > 10 )
			{
				Process.Start(urlServicio);
			}

			//Deseleccionar todos cuando se logre imprimir
			string squery = String.Format("EXEC EncuestasSelQRDeshabilitar {0}", idBodg);
			Funcion.EjecutaSQL(cdsSeteoF,squery);

			MessageBox.Show("Será redirigido para poder imprimir el código QR\n\npor favor espere", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

			idBodg = (int)this.cmbBodega.Value;
			string ugSQL = string.Format("Exec EncuestasSelQR {0}",idBodg);
			ugHabitaciones.DataSource = Funcion.dvProcedimiento(cdsSeteoF, ugSQL);

		}
	}
}
