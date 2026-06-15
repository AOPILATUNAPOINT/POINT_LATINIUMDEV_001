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
	/// Descripción breve de frmEmpresasExternas.
	/// </summary>
	public class frmEmpresasExternas : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDocumentos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Button btnCancelar;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmEmpresasExternas()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCobranzaExterna");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Empresa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaApertura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaCierre");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCobranzaExterna");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Empresa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaApertura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaCierre");
			this.uGridDocumentos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.uGridDocumentos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
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
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 250;
			ultraGridColumn3.Header.Caption = "Fecha Apertura";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn4.Header.Caption = "Fecha Cierre";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridDocumentos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridDocumentos.DisplayLayout.GroupByBox.Hidden = true;
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDocumentos.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.uGridDocumentos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridDocumentos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridDocumentos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.uGridDocumentos.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDocumentos.DisplayLayout.Override.CellAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 8F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDocumentos.DisplayLayout.Override.HeaderAppearance = appearance5;
			this.uGridDocumentos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance6.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance6.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.RowAlternateAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.uGridDocumentos.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridDocumentos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDocumentos.Location = new System.Drawing.Point(0, 0);
			this.uGridDocumentos.Name = "uGridDocumentos";
			this.uGridDocumentos.Size = new System.Drawing.Size(456, 288);
			this.uGridDocumentos.TabIndex = 765;
			this.uGridDocumentos.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridDocumentos_AfterRowInsert);
			this.uGridDocumentos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridDocumentos_KeyDown);
			this.uGridDocumentos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDocumentos_AfterCellUpdate);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(System.DateTime);
			ultraDataColumn4.DataType = typeof(System.DateTime);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.Location = new System.Drawing.Point(8, 296);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(88, 23);
			this.btnAceptar.TabIndex = 766;
			this.btnAceptar.Text = "&Guardar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.Location = new System.Drawing.Point(104, 296);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(88, 23);
			this.btnCancelar.TabIndex = 767;
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
			// frmEmpresasExternas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(456, 326);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.uGridDocumentos);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmEmpresasExternas";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Empresas Externas";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmEmpresasExternas_KeyDown);
			this.Load += new System.EventHandler(this.frmEmpresasExternas_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridDocumentos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmEmpresasExternas_Load(object sender, System.EventArgs e)
		{
			string sSQLPa = string.Format(" SELECT idCobranzaExterna, Empresa,FechaApertura, FechaCierre  FROM CobranzaExterna");
			FuncionesProcedimientos.dsGeneral(sSQLPa, this.uGridDocumentos);

		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void uGridDocumentos_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idCobranzaExterna"].Value = 0;
			e.Row.Cells["FechaApertura"].Value = DateTime.Today  ;	
		}

		private void uGridDocumentos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender,e,this.uGridDocumentos);
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void frmEmpresasExternas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape)  this.DialogResult = DialogResult.No;	
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
     DateTime dtFecha =  DateTime.Today  ;	
     
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)
			{	

				
				if (dr.Cells["FechaApertura"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese la fecha de Apertura", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridDocumentos.ActiveRow = dr;
					this.uGridDocumentos.ActiveRow.Selected = true;
					return;
				}

				if ( (int)dr.Cells["idCobranzaExterna"].Value == 0)
				{
					if (Convert.ToDateTime(dr.Cells["FechaApertura"].Value) < Convert.ToDateTime(dtFecha))
					{
						MessageBox.Show("La Fecha de Apertura No Puede Ser Menor A La fecha Actual ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridDocumentos.ActiveRow = dr;
						this.uGridDocumentos.ActiveRow.Selected = true;
						return;
					}	
				}
				
				if ( (int)dr.Cells["idCobranzaExterna"].Value == 0)
				{
					if (Convert.ToDateTime(dr.Cells["FechaApertura"].Value) > Convert.ToDateTime(dtFecha))
					{
						MessageBox.Show("La Fecha de Apertura No Puede Ser Mayor A La fecha Actual ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridDocumentos.ActiveRow = dr;
						this.uGridDocumentos.ActiveRow.Selected = true;
						return;
					}	
				}

				if (dr.Cells["Empresa"].Value.ToString().Length == 0)
				{
					MessageBox.Show("Escriba el Detalle ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridDocumentos.ActiveRow = dr;
					this.uGridDocumentos.ActiveRow.Selected = true;
					return;
				}


				if (dr.Cells["FechaCierre"].Value!= System.DBNull.Value)
				{
					if (Convert.ToDateTime(dr.Cells["FechaCierre"].Value) > Convert.ToDateTime("01/01/2000"))
					{

						if (dr.Cells["FechaApertura"].Value != System.DBNull.Value)
						{
							if (Convert.ToDateTime(dr.Cells["FechaCierre"].Value) < Convert.ToDateTime(dr.Cells["FechaApertura"].Value))
							{
								MessageBox.Show("La Fecha de Cierre No Puede Ser Menor A La De Apertura ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
								this.uGridDocumentos.ActiveRow = dr;
								this.uGridDocumentos.ActiveRow.Selected = true;
								return;
							}	
							if (Convert.ToDateTime(dr.Cells["FechaCierre"].Value) > Convert.ToDateTime(dtFecha))
							{
								MessageBox.Show("La Fecha de Cierre No Puede Ser Mayor A La Actual", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
								this.uGridDocumentos.ActiveRow = dr;
								this.uGridDocumentos.ActiveRow.Selected = true;
								return;
							}	
						}
					}
				}
				
			}
			
			if (DialogResult.Yes == MessageBox.Show(string.Format(" Desea Guardar Cambios " , 0), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
					 
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
							
								DateTime dtfechacierre = DateTime.Parse("01/01/2000");
								if (dr.Cells["FechaCierre"].Value!= System.DBNull.Value)dtfechacierre = Convert.ToDateTime(dr.Cells["FechaCierre"].Value);

								string  sSQLDet = string.Format(" EXEC  GuardaEmpresaExterna {0}, '{1}', '{2}', '{3}' ", 
									(int)dr.Cells["idCobranzaExterna"].Value , Convert.ToString( dr.Cells["Empresa"].Value.ToString()),
									Convert.ToDateTime(dr.Cells["FechaApertura"].Value).ToString("yyyyMMdd"), Convert.ToDateTime(dtfechacierre).ToString("yyyyMMdd"));
								oCmdActualiza.CommandText = sSQLDet;
								oCmdActualiza.ExecuteNonQuery();	
							
						}
						oTransaction.Commit();

						MessageBox.Show("Registro Guardado Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);

					
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
			else
			{
				this.DialogResult = DialogResult.No;	

			}
		}
    	bool bActualiza = false;
		private void uGridDocumentos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bActualiza) return;
			if (e.Cell.Column.ToString() == "FechaApertura")
			{	
				bActualiza = true;

				if ((int)  uGridDocumentos.ActiveRow.Cells["idCobranzaExterna"].Value > 0)
				{										
						e.Cell.Row.Cells["FechaApertura"].Value = (DateTime)e.Cell.OriginalValue ;								
				}
				bActualiza = false;
			}

			if (e.Cell.Column.ToString() == "FechaCierre")
			{	
				bActualiza = true;

				if ((int)  uGridDocumentos.ActiveRow.Cells["idCobranzaExterna"].Value > 0)
				{					
					if (DialogResult.No == MessageBox.Show("¿Esta seguro de Modificar la Fecha De Cierre ?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
					{						
						e.Cell.Row.Cells["FechaCierre"].Value = "01/01/2000";				
					}				
				}
				bActualiza = false;
			}
		}	
	}
}
