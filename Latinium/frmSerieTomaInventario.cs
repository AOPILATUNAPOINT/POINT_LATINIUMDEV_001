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
	/// Descripción breve de frmSerieTomaInventario.
	/// </summary>
	public class frmSerieTomaInventario : System.Windows.Forms.Form
	{
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmSerieTomaInventario()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		int iCierreConteo = 0;
		int iBodega = 0;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Button btnCancelar;
		private C1.Data.C1DataSet cdsSeteoF;
		int iArticulo = 0;
    string sCodigo ="";
    int bExibido = 0;
		DateTime dtFechas ;
   int idCierreConteoInventario = 0;
		public frmSerieTomaInventario(int ICierreConteo , int IBodega , int IArticulo, string SCodigo , int BExibido, DateTime DtFechas, int IdCierreConteoInventario )
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			 iCierreConteo = ICierreConteo;
			 iBodega = IBodega;
			 iArticulo = IArticulo;
       sCodigo = SCodigo ;
			dtFechas =DtFechas;
		  bExibido= BExibido;
     idCierreConteoInventario = IdCierreConteoInventario;
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnAgregar", 0);
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmSerieTomaInventario));
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSerial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serial");
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			appearance1.ForeColor = System.Drawing.Color.Black;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 225;
			ultraGridColumn2.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Width = 345;
			ultraGridColumn3.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance2.Image = ((object)(resources.GetObject("appearance2.Image")));
			ultraGridColumn3.CellAppearance = appearance2;
			appearance3.Image = ((object)(resources.GetObject("appearance3.Image")));
			appearance3.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn3.CellButtonAppearance = appearance3;
			ultraGridColumn3.Header.Caption = "...";
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn3.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn3.Width = 34;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			ultraGridBand1.Header.Caption = "SERIALES";
			ultraGridBand1.HeaderVisible = true;
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance4;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance5.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 8F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance6;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance7.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance7.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 8);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(400, 248);
			this.ultraGrid1.TabIndex = 6;
			this.ultraGrid1.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_ClickCellButton);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.Location = new System.Drawing.Point(8, 264);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(88, 23);
			this.btnAceptar.TabIndex = 7;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.Location = new System.Drawing.Point(104, 264);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(88, 23);
			this.btnCancelar.TabIndex = 8;
			this.btnCancelar.Text = "&Cancelar";
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
			// frmSerieTomaInventario
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(416, 294);
			this.ControlBox = false;
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.ultraGrid1);
			this.Name = "frmSerieTomaInventario";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = ".";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmSerieTomaInventario_KeyDown);
			this.Load += new System.EventHandler(this.frmSerieTomaInventario_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmSerieTomaInventario_Load(object sender, System.EventArgs e)
		{
			this.Text  = sCodigo;
			string sSQL = string.Format("Exec ReporteSerialesInventario {0}, {1}", 
				 iBodega, iArticulo);
			FuncionesProcedimientos.dsGeneral(sSQL, this.ultraGrid1);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
	
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			
			this.DialogResult = DialogResult.OK;
	
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
	
		}

		private void frmSerieTomaInventario_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		
		}

		private void ultraGrid1_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "btnAgregar")
			{
				DateTime dtFecha = dtFechas;
	
				string sMensaje = "";
        int idConteoInventario = 0;
				int iRes = 0;
				int idSerial = 0;
				//int idArticulo = 0;

				#region Valida En Transacciones
				int idTipoFactura = 0;
				string sTipoFactura = "";
				string sNumero = "";
				string sSQLVT = string.Format("Exec ConteoValidaSerialEnTransacciones '{0}', {1}", e.Cell.Row.Cells["Serial"].Value.ToString(), iBodega);
				SqlDataReader drVT = Funcion.rEscalarSQL(cdsSeteoF, sSQLVT, true);
				drVT.Read();
				if (drVT.HasRows)
				{
					idTipoFactura = drVT.GetInt32(0);
					sTipoFactura = drVT.GetString(1);
					sNumero = drVT.GetString(2);
					drVT.Close();
				}
				drVT.Close();

				if (idTipoFactura == 2)
				{
					MessageBox.Show(string.Format("Esta Serial esta registrada en el Documento '{0}' Numero '{1}', No puede registrarla", sTipoFactura, sNumero), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);	
					
					return;
				}
				else if (idTipoFactura > 0 && DialogResult.No == MessageBox.Show(
					string.Format("Esta Serial esta registrada en el Documento '{0}' Numero '{1}',\n\n\n¿Seguro de Ingresarla?", sTipoFactura, sNumero), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
					return;
				#endregion Valida En Transacciones

				#region Valida Con Articulo Seleccionado
				bool bManejaSerial = false;
				if (iArticulo > 0) 
				{
					bManejaSerial = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select ManejaSerial From Articulo Where idArticulo = {0}", iArticulo), true);
					
			
					if (bManejaSerial)
					{
						#region Maneja Serial
						sMensaje = "";
						string sSQL = string.Format("Exec ValidaSerialEnConteo {0}, '{1}', '{2}', {3}", iArticulo, e.Cell.Row.Cells["Serial"].Value.ToString(), dtFecha.ToString("yyyyMMdd"), iBodega);
						sMensaje = Funcion.sEscalarSQL(cdsSeteoF, sSQL, true);

						if (sMensaje.Length > 0)
						{
							MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							return;
						}
				
						sMensaje = "";
						string sSQLSerial = string.Format("Exec ValidaSerialExistenteEnSeriales {0}, '{1}'", iArticulo, e.Cell.Row.Cells["Serial"].Value.ToString());
						SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQLSerial, true);
						dr.Read();
						if (dr.HasRows)
						{
							iRes = dr.GetInt32(0);
							sMensaje = dr.GetString(1);
							idSerial = dr.GetInt32(2);
						}
						dr.Close();
				
						if (iRes == 1) // SERIAL NO EXISTE
						{
							MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						
							return;
						}

						if (iRes == 2) // SERIAL NO EXISTE PERO REGISTRA CON EL ARTICULO SELECCIONADO
						{
							if (DialogResult.Yes == MessageBox.Show(sMensaje + " " + sCodigo, "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
							{
								string sSQLGC = string.Format("Exec GuardaConteoInventario {0}, '{1}', {2}, {3}, '{4}', {5}, {6}, {7}, {8}, {9}",
									iArticulo, dtFecha.ToString("yyyyMMdd HH:mm"), iBodega, idSerial, e.Cell.Row.Cells["Serial"].Value.ToString(), 1, 1, 0,bExibido, 1);
//								Funcion.EjecutaSQL(cdsSeteoF, sSQLGC, true);

								idConteoInventario = Funcion.iEscalarSQL(cdsSeteoF, sSQLGC, true);

								string sSQLGCs = string.Format("Exec [GuardaConsultaSerieConteo] {0}, {1}, {2}, {3}, {4}",
									iArticulo,  iBodega, idConteoInventario, idCierreConteoInventario, idSerial );
										Funcion.EjecutaSQL(cdsSeteoF, sSQLGCs, true);

								string sSQLc = string.Format("Exec ReporteSerialesInventario {0}, {1}", 
									iBodega, iArticulo);
								FuncionesProcedimientos.dsGeneral(sSQLc, this.ultraGrid1);

								return;
							}
						}

						if (iRes == 3) // SERIAL SI EXISTE CAMBIA ARTICULO ACTUAL DE LA SERIAL POR EL SELECCIONADO 
						{
							if (DialogResult.Yes == MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
							{
								string sSQLGC = string.Format("Exec GuardaConteoInventario {0}, '{1}', {2}, {3}, '{4}', {5}, {6}, {7}, {8}, {9}",
								iArticulo, dtFecha.ToString("yyyyMMdd HH:mm"), iBodega, idSerial, e.Cell.Row.Cells["Serial"].Value.ToString(), 1, 1, 1,bExibido, 1);
//								Funcion.EjecutaSQL(cdsSeteoF, sSQLGC, true);
								idConteoInventario = Funcion.iEscalarSQL(cdsSeteoF, sSQLGC, true);

								
								string sSQLGCs = string.Format("Exec [GuardaConsultaSerieConteo] {0}, {1}, {2}, {3}, {4}",
									iArticulo,  iBodega, idConteoInventario, idCierreConteoInventario, idSerial );
								Funcion.EjecutaSQL(cdsSeteoF, sSQLGCs, true);

								string sSQLa = string.Format("Exec ReporteSerialesInventario {0}, {1}", 
									iBodega, iArticulo);
								FuncionesProcedimientos.dsGeneral(sSQLa, this.ultraGrid1);

								return;
							}
						}

						if (iRes == 4) // SERIAL SI EXISTE MANTIENE ARTICULO ACTUAL
						{
							string sSQLGC = string.Format("Exec GuardaConteoInventario {0}, '{1}', {2}, {3}, '{4}', {5}, {6}, {7}, {8}, {9}",
								iArticulo, dtFecha.ToString("yyyyMMdd HH:mm"), iBodega, idSerial, e.Cell.Row.Cells["Serial"].Value.ToString(), 1, 1, 0,bExibido, 1);
//							Funcion.EjecutaSQL(cdsSeteoF, sSQLGC, true);
							idConteoInventario = Funcion.iEscalarSQL(cdsSeteoF, sSQLGC, true);

							
							string sSQLGCs = string.Format("Exec [GuardaConsultaSerieConteo] {0}, {1}, {2}, {3}, {4}",
								iArticulo,  iBodega, idConteoInventario, idCierreConteoInventario, idSerial );
							Funcion.EjecutaSQL(cdsSeteoF, sSQLGCs, true);


							string sSQLb = string.Format("Exec ReporteSerialesInventario {0}, {1}", 
								iBodega, iArticulo);
							FuncionesProcedimientos.dsGeneral(sSQLb, this.ultraGrid1);

							return;						
						}
						#endregion Maneja Serial
					}
				}
				#endregion Valida Con Articulo Seleccionado
			}
		}
	}
}
