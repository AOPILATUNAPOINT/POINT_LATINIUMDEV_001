using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using Infragistics.Win.UltraWinGrid;
using System.Globalization;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmTurnosMercaderia.
	/// </summary>
	public class frmTurnosMercaderia : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Acceso miAcceso;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridTurnos;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmTurnosMercaderia()
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Hora");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmTurnosMercaderia));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label6 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.btnVer = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.uGridTurnos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridTurnos)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraDataSource1
			// 
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// label2
			// 
			this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label2.BackColor = System.Drawing.Color.Red;
			this.label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label2.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.ForeColor = System.Drawing.Color.White;
			this.label2.Location = new System.Drawing.Point(1084, 58);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(200, 23);
			this.label2.TabIndex = 273;
			this.label2.Text = "Ordenes de Compra";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label3
			// 
			this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label3.BackColor = System.Drawing.Color.White;
			this.label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label3.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.ForeColor = System.Drawing.Color.Black;
			this.label3.Location = new System.Drawing.Point(1084, 8);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(200, 23);
			this.label3.TabIndex = 274;
			this.label3.Text = "Disponible";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
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
			// label6
			// 
			this.label6.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label6.BackColor = System.Drawing.Color.Blue;
			this.label6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label6.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label6.ForeColor = System.Drawing.Color.White;
			this.label6.Location = new System.Drawing.Point(1084, 33);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(200, 23);
			this.label6.TabIndex = 277;
			this.label6.Text = "No Disponible";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label8
			// 
			this.label8.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label8.BackColor = System.Drawing.Color.Yellow;
			this.label8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label8.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.ForeColor = System.Drawing.Color.Black;
			this.label8.Location = new System.Drawing.Point(1084, 83);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(200, 23);
			this.label8.TabIndex = 279;
			this.label8.Text = "Ordenes de Producción";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(8, 8);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(100, 28);
			this.btnVer.TabIndex = 741;
			this.btnVer.Text = "&Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 120);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1292, 8);
			this.groupBox1.TabIndex = 742;
			this.groupBox1.TabStop = false;
			// 
			// uGridTurnos
			// 
			this.uGridTurnos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridTurnos.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridTurnos.DisplayLayout.Appearance = appearance1;
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridTurnos.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.uGridTurnos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridTurnos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.uGridTurnos.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridTurnos.DisplayLayout.Override.CellAppearance = appearance4;
			this.uGridTurnos.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 9F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridTurnos.DisplayLayout.Override.HeaderAppearance = appearance5;
			appearance6.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance6.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTurnos.DisplayLayout.Override.RowAlternateAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTurnos.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTurnos.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.uGridTurnos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridTurnos.Location = new System.Drawing.Point(8, 136);
			this.uGridTurnos.Name = "uGridTurnos";
			this.uGridTurnos.Size = new System.Drawing.Size(1276, 272);
			this.uGridTurnos.TabIndex = 743;
			this.uGridTurnos.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridTurnos_ClickCellButton);
			this.uGridTurnos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridTurnos_InitializeLayout);
			// 
			// frmTurnosMercaderia
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1292, 416);
			this.Controls.Add(this.uGridTurnos);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.KeyPreview = true;
			this.Name = "frmTurnosMercaderia";
			this.Text = "Turnos de Recepción de Mercaderia";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmTurnosMercaderia_KeyDown);
			this.Load += new System.EventHandler(this.frmTurnosMercaderia_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridTurnos)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmTurnosMercaderia_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0322");

			if (!Funcion.Permiso("881", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a turnos de recepcion de mercaderia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

				BeginInvoke(new MethodInvoker(this.Close));				

				return;
			}

			Consulta();
		}

		private void Consulta()
		{
			this.uGridTurnos.DataSource = null;
			
			this.uGridTurnos.DataSource = FuncionesProcedimientos.dtGeneral("Exec ConsultaTurnosOrdenes");

			DiseñoGridTurno();
		}

		private void DiseñoGridTurno()
		{
			try
			{				
				this.uGridTurnos.DisplayLayout.Bands[0].Columns["idTurno"].Hidden = true;
				this.uGridTurnos.DisplayLayout.Bands[0].Columns["Fecha"].Header.Caption = "Fecha";
				this.uGridTurnos.DisplayLayout.Bands[0].Columns["Fecha"].Width = 75;
				this.uGridTurnos.DisplayLayout.Bands[0].Columns["Fecha"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
				this.uGridTurnos.DisplayLayout.Bands[0].Columns["DiaT"].Header.Caption = "Día";
				this.uGridTurnos.DisplayLayout.Bands[0].Columns["DiaT"].Width = 80;
				this.uGridTurnos.DisplayLayout.Bands[0].Columns["DiaT"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;

				if (this.uGridTurnos.Rows.Count > 0)
				{
					int iContColumnas = (int)this.uGridTurnos.DisplayLayout.Bands[0].Columns.Count;					
					
					string nombrCol = "";

					for (int j = 3; j < iContColumnas; j++)
					{
						nombrCol = this.uGridTurnos.DisplayLayout.Bands[0].Columns[j].Header.Caption;	
						this.uGridTurnos.DisplayLayout.Bands[0].Columns[nombrCol].Style = ColumnStyle.Button;
						this.uGridTurnos.DisplayLayout.Bands[0].Columns[nombrCol].Width = 75;
						this.uGridTurnos.DisplayLayout.Bands[0].Columns[nombrCol].Header.Caption = nombrCol;	
					}
					
					string sColumna = "";

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridTurnos.Rows.All)
					{
						for (int i = 3; i < iContColumnas; i++)
						{
							sColumna = this.uGridTurnos.DisplayLayout.Bands[0].Columns[i].Header.Caption;							
						
							if (dr.Cells[sColumna].Value.ToString().Trim().Length == 0) 
							{
								dr.Cells[sColumna].Appearance.BackColor = Color.White;
								dr.Cells[sColumna].Appearance.BackColor2 = Color.White;								
							}
							else if (dr.Cells[sColumna].Value.ToString().Trim().Length > 0) 
							{
								string sNumero = dr.Cells[sColumna].Value.ToString();
								
								if(sNumero.ToString().Substring(0, 2) == "OC")
								{
									dr.Cells[sColumna].Appearance.BackColor = Color.Red;
									dr.Cells[sColumna].Appearance.BackColor2 = Color.Red;
								}
								else if (sNumero.ToString().Substring(0, 2) == "OP")
								{
									dr.Cells[sColumna].Appearance.BackColor = Color.Yellow; 
									dr.Cells[sColumna].Appearance.BackColor2 = Color.Yellow; 
								}								
							}

							if ((sColumna == "T0700" || sColumna == "T0730" || sColumna == "T1530" || sColumna == "T1600" || sColumna == "T1630" || sColumna == "T1700" || sColumna == "T1730" || sColumna == "T1800" || sColumna == "T1830" || sColumna == "T1900") && dr.Cells[sColumna].Value.ToString().Trim().Length == 0)
							{								
								dr.Cells[sColumna].Appearance.BackColor = Color.Blue;
								dr.Cells[sColumna].Appearance.BackColor2 = Color.Blue;
							}							
						}
					}
				}			
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			Consulta();			
		}

		private void uGridTurnos_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{			
			string sColumna = e.Cell.Column.Key;
			
			if (e.Cell.Column.ToString() == sColumna)
			{
				if (e.Cell.Row.Cells[sColumna].Appearance.BackColor == Color.Blue)
				{
					MessageBox.Show("Horario No Disponible", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if (e.Cell.Row.Cells[sColumna].Value.ToString().Trim().Length == 0)
				{
					if (!miAcceso.Nuevo)
					{
						MessageBox.Show("No tiene acceso a generar turnos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}

					using (frmbotones Opciones = new frmbotones())
					{
						if (DialogResult.OK == Opciones.ShowDialog())
						{
							string sFecha = e.Cell.Row.Cells["Fecha"].Value.ToString();

							string sHorallegada = sColumna.Substring(1, 2);
							
							DateTime dtHLLegada = DateTime.Today;
							
							dtHLLegada = DateTime.Parse(sFecha + " " + sHorallegada + ":" + sColumna.Substring(3,2));

							if (Opciones.iOrigen == 1)
							{																
								frmOrdenesDeCompra OCB = new frmOrdenesDeCompra(0, 1, dtHLLegada, sColumna);
								OCB.Show();
							}
							else if (Opciones.iOrigen == 2)
							{
								frmGeneraOrdenesDeProduccion OP = new frmGeneraOrdenesDeProduccion(0, 2, dtHLLegada, sColumna);
								OP.Show();
							}
						}

						return;
					}
				}
				else if (e.Cell.Row.Cells[sColumna].Value.ToString().Trim().Length > 0)
				{
					if (!miAcceso.Buscar)
					{
						MessageBox.Show("No tiene acceso a ver los documentos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}

					string sNumero = e.Cell.Row.Cells[sColumna].Value.ToString();
					
					int idDocumento = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ConsultaIdDocumentoTurnosMercaderia '{0}', {1}", sColumna, (int)e.Cell.Row.Cells["idTurno"].Value));
	
					if (sNumero.ToString().Substring(0, 2) == "OC")
					{						
						frmOrdenesDeCompra OC = new frmOrdenesDeCompra(idDocumento);
						OC.Show();						
					}
					else if (sNumero.ToString().Substring(0, 2) == "OP")
					{
						frmGeneraOrdenesDeProduccion OP = new frmGeneraOrdenesDeProduccion(idDocumento);
						OP.ShowDialog();
					}
				}
			}
		}		

		private void frmTurnosMercaderia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void uGridTurnos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}		
	}
}
