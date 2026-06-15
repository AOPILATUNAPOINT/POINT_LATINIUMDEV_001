using System;
using System.Globalization;
using System.Data;
using System.Data.SqlClient;
using C1.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using System.Web.Mail;
using CrystalDecisions.Shared;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmIncentivosVentaGarantia.
	/// </summary>
	public class frmIncentivosVentaGarantia : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource13;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnSalir;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmIncentivosVentaGarantia()
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idIncentivosVentaGarantia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Desde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Hasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MontoDesde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MontoHasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ganancia");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idIncentivosVentaGarantia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desde");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hasta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MontoDesde");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MontoHasta");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ganancia");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmIncentivosVentaGarantia));
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraDataSource13 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridInformacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource13)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).BeginInit();
			this.SuspendLayout();
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
			// ultraDataSource13
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn1.DefaultValue = 0;
			ultraDataColumn1.Tag = typeof(int);
			ultraDataColumn2.DataType = typeof(System.DateTime);
			ultraDataColumn3.DataType = typeof(System.DateTime);
			ultraDataColumn4.DataType = typeof(System.Decimal);
			ultraDataColumn5.DataType = typeof(System.Decimal);
			ultraDataColumn6.DataType = typeof(System.Decimal);
			this.ultraDataSource13.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn1,
																																	ultraDataColumn2,
																																	ultraDataColumn3,
																																	ultraDataColumn4,
																																	ultraDataColumn5,
																																	ultraDataColumn6});
			this.ultraDataSource13.CellDataRequested += new Infragistics.Win.UltraWinDataSource.CellDataRequestedEventHandler(this.ultraDataSource13_CellDataRequested);
			// 
			// uGridInformacion
			// 
			this.uGridInformacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridInformacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridInformacion.DataSource = this.ultraDataSource13;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridInformacion.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Format = "";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn3.Format = "";
			ultraGridColumn3.Header.VisiblePosition = 2;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance2;
			ultraGridColumn4.Format = "#,##0.00";
			ultraGridColumn4.Header.Caption = "Monto Desde";
			ultraGridColumn4.Header.VisiblePosition = 3;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance3;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.Caption = "Monto Hasta";
			ultraGridColumn5.Header.VisiblePosition = 4;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance4;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.uGridInformacion.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridInformacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridInformacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridInformacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance5.BackColor = System.Drawing.Color.Transparent;
			this.uGridInformacion.DisplayLayout.Override.CardAreaAppearance = appearance5;
			appearance6.ForeColor = System.Drawing.Color.Black;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInformacion.DisplayLayout.Override.CellAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance7.FontData.BoldAsString = "True";
			appearance7.FontData.Name = "Arial";
			appearance7.FontData.SizeInPoints = 8F;
			appearance7.ForeColor = System.Drawing.Color.White;
			appearance7.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInformacion.DisplayLayout.Override.HeaderAppearance = appearance7;
			this.uGridInformacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.uGridInformacion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInformacion.Location = new System.Drawing.Point(8, 8);
			this.uGridInformacion.Name = "uGridInformacion";
			this.uGridInformacion.Size = new System.Drawing.Size(464, 184);
			this.uGridInformacion.TabIndex = 1002;
			this.uGridInformacion.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridInformacion_AfterRowInsert);
			this.uGridInformacion.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridInformacion_InitializeLayout);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(16, 208);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(72, 23);
			this.btnGuardar.TabIndex = 1005;
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
			this.btnSalir.Location = new System.Drawing.Point(104, 208);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(72, 23);
			this.btnSalir.TabIndex = 1006;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// frmIncentivosVentaGarantia
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(480, 238);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.uGridInformacion);
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmIncentivosVentaGarantia";
			this.Text = "Freelance Venta Garantía";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmIncentivosVentaGarantia_KeyDown);
			this.Load += new System.EventHandler(this.frmIncentivosVentaGarantia_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource13)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}

		private void frmIncentivosVentaGarantia_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "280101");

			if (!Funcion.Permiso("2100", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a configuración de Freelance.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			string sSQLPa = string.Format(" Exec ConsultarIncentivosVentaGarantia   ");
			FuncionesProcedimientos.dsGeneralModulo(sSQLPa, this.uGridInformacion);
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close ();
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if((int)this.uGridInformacion.Rows.Count == 0)
			{
				MessageBox.Show("Ingrese al menos un registro.", "Point Technolody", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.uGridInformacion.Focus();
				return;
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInformacion.Rows.All)
			{
				#region validarFechas

				if (dr.Cells["MontoDesde"].Value == System.DBNull.Value)
				{
					MessageBox.Show("El campo Monto Desde no puede estar vacio. ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridInformacion.ActiveRow = dr;
					this.uGridInformacion.ActiveRow.Selected = true;

					return;
				}
				if (dr.Cells["MontoHasta"].Value == System.DBNull.Value)
				{
					MessageBox.Show("El campo Monto Hasta no puede estar vacio. ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridInformacion.ActiveRow = dr;
					this.uGridInformacion.ActiveRow.Selected = true;

					return;
				}
				if (dr.Cells["Ganancia"].Value == System.DBNull.Value)
				{
					MessageBox.Show("El campo Ganancia no puede estar vacio. ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridInformacion.ActiveRow = dr;
					this.uGridInformacion.ActiveRow.Selected = true;

					return;
				}


				if(this.uGridInformacion.Rows.Count == 0)
				{
					MessageBox.Show("Ingrese los datos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridInformacion.ActiveRow = dr;
					this.uGridInformacion.ActiveRow.Selected = true;

					return;
				}

				if(dr.Cells["Desde"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese la fecha inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridInformacion.ActiveRow = dr;
					this.uGridInformacion.ActiveRow.Selected = true;

					return;
				}

				if(dr.Cells["Hasta"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese la fecha final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridInformacion.ActiveRow = dr;
					this.uGridInformacion.ActiveRow.Selected = true;

					return;
				}

				
				if(Convert.ToDateTime( dr.Cells["Desde"].Value ) > Convert.ToDateTime( dr.Cells["Hasta"].Value ) )
				{
					MessageBox.Show("La fecha inicial no puede ser mayor a la fecha final.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridInformacion.ActiveRow = dr;
					this.uGridInformacion.ActiveRow.Selected = true;

					return;
				}
        

				if((decimal)dr.Cells["MontoDesde"].Value == 0)
				{
					MessageBox.Show("Ingrese el valor Desde.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridInformacion.ActiveRow = dr;
					this.uGridInformacion.ActiveRow.Selected = true;

					return;
				}
				if ((decimal)dr.Cells["MontoDesde"].Value > (decimal)dr.Cells["MontoHasta"].Value)
				{
					MessageBox.Show("Monto Desde no puede ser mayor al Monto Hasta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridInformacion.ActiveRow = dr;
					this.uGridInformacion.ActiveRow.Selected = true;

					return;
				}

				if((decimal)dr.Cells["MontoHasta"].Value == 0)
				{
					MessageBox.Show("Ingrese el valor Hasta.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridInformacion.ActiveRow = dr;
					this.uGridInformacion.ActiveRow.Selected = true;

					return;
				}

				if((decimal)dr.Cells["Ganancia"].Value == 0)
				{
					MessageBox.Show("Ingrese el valor de la Ganancia.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridInformacion.ActiveRow = dr;
					this.uGridInformacion.ActiveRow.Selected = true;

					return;
				}


				#endregion validarFechas
			}
			#region Guardar

			//			DateTime dtInicia = DateTime.Today;
			//			DateTime dtRetorna = DateTime.Today;
			//				dtInicia = (DateTime)this.dtInicia.Value;
			//				dtRetorna = (DateTime)this.dtRetorna.Value;

			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + "; max pool size=1024"))
			{
				oConexion.Open();

				SqlTransaction oTransaction = null;
				SqlCommand oCmdActualiza = oConexion.CreateCommand();
				oCmdActualiza.Connection = oConexion;
				oCmdActualiza.CommandType = CommandType.Text;
				oCmdActualiza.CommandTimeout = 120;

				try
				{
					oTransaction = oConexion.BeginTransaction();
					oCmdActualiza.Transaction = oTransaction;


					#region Graba DetalleMaestro

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInformacion.Rows.All)
					{
						string  sSQLDet = string.Format("Exec GuardarIncentivosVentaGarantia {0}, '{1}', '{2}', {3}, {4}, {5}" ,

							Convert.ToInt32 ( dr.Cells["idIncentivosVentaGarantia"].Value), 
							Convert.ToDateTime(dr.Cells["Desde"].Value).ToString("yyyyMMdd"),
							Convert.ToDateTime(dr.Cells["Hasta"].Value).ToString("yyyyMMdd"),
							Convert.ToDecimal ( dr.Cells["MontoDesde"].Value),
							Convert.ToDecimal ( dr.Cells["MontoHasta"].Value),
							Convert.ToDecimal ( dr.Cells["Ganancia"].Value)
							);
						oCmdActualiza.CommandText = sSQLDet;
						oCmdActualiza.ExecuteNonQuery();	
					}			
					#endregion Graba DetalleMaestro
					
          
					oTransaction.Commit();

					Cursor = Cursors.Default;	

					string sSQLPa = string.Format(" Exec ConsultarIncentivosVentaGarantia   ");
					FuncionesProcedimientos.dsGeneralModulo(sSQLPa, this.uGridInformacion);

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
				
			#endregion Guardar
		}

		private void uGridInformacion_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "MontoDesde", 2);
			FuncionesProcedimientos.FormatoGrid(e, "MontoHasta", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Ganancia", 2);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		
		}

		private void ultraDataSource13_CellDataRequested(object sender, Infragistics.Win.UltraWinDataSource.CellDataRequestedEventArgs e)
		{
		
		}

		private void uGridInformacion_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idIncentivosVentaGarantia"].Value = 0;
			e.Row.Cells["MontoDesde"].Value = 0.00;
			e.Row.Cells["MontoHasta"].Value = 0.00;
			e.Row.Cells["Ganancia"].Value = 0.00;
		}

		private void frmIncentivosVentaGarantia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();		
		}
	}
}
