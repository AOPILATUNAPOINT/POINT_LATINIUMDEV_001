using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmContacto.
	/// </summary>
	public class frmContacto : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnAceptar;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridContacto;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmContacto()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmContacto));
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idContacto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Contacto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Empresa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Convencional");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Extension");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Celular");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EMail");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notas");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idContacto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Contacto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Empresa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Convencional");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Extension");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EMail");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridContacto = new Infragistics.Win.UltraWinGrid.UltraGrid();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridContacto)).BeginInit();
			this.SuspendLayout();
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnCancelar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(104, 272);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 430;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnAceptar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnAceptar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnAceptar.Image = ((System.Drawing.Image)(resources.GetObject("btnAceptar.Image")));
			this.btnAceptar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAceptar.Location = new System.Drawing.Point(8, 272);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(76, 23);
			this.btnAceptar.TabIndex = 431;
			this.btnAceptar.Text = "Aceptar";
			this.btnAceptar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
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
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
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
			// uGridContacto
			// 
			this.uGridContacto.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridContacto.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridContacto.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridContacto.DisplayLayout.Appearance = appearance1;
			this.uGridContacto.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "RetDetalle";
			ultraGridColumn1.Header.VisiblePosition = 7;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 100;
			ultraGridColumn2.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn2.Width = 190;
			ultraGridColumn3.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn3.Width = 221;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance2;
			ultraGridColumn4.Header.Caption = "Teléfono";
			ultraGridColumn4.Header.VisiblePosition = 2;
			ultraGridColumn4.Width = 87;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance3;
			ultraGridColumn5.Header.Caption = "Ext.";
			ultraGridColumn5.Header.VisiblePosition = 3;
			ultraGridColumn5.Width = 93;
			ultraGridColumn6.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance4;
			ultraGridColumn6.Header.VisiblePosition = 4;
			ultraGridColumn6.Width = 93;
			ultraGridColumn7.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn7.Header.Caption = "E-Mail";
			ultraGridColumn7.Header.VisiblePosition = 5;
			ultraGridColumn7.Width = 121;
			ultraGridColumn8.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn8.Header.VisiblePosition = 6;
			ultraGridColumn8.Width = 214;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.uGridContacto.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridContacto.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridContacto.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridContacto.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance5.BackColor = System.Drawing.Color.Transparent;
			this.uGridContacto.DisplayLayout.Override.CardAreaAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 10F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridContacto.DisplayLayout.Override.HeaderAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridContacto.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridContacto.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.uGridContacto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridContacto.Location = new System.Drawing.Point(8, 8);
			this.uGridContacto.Name = "uGridContacto";
			this.uGridContacto.Size = new System.Drawing.Size(1040, 256);
			this.uGridContacto.TabIndex = 432;
			this.uGridContacto.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridContacto_AfterRowInsert);
			this.uGridContacto.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.uGridContacto_KeyPress);
			this.uGridContacto.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridContacto_KeyDown);
			// 
			// frmContacto
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1048, 302);
			this.Controls.Add(this.uGridContacto);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnCancelar);
			this.KeyPreview = true;
			this.Name = "frmContacto";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Contacto";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmContacto_KeyDown);
			this.Load += new System.EventHandler(this.frmContacto_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridContacto)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
  
	bool 	bEdicion = true;
		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			#region validar campos datagrid 
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridContacto.Rows.All)
			{
				if (dr.Cells["Contacto"].Value.ToString().Trim().Length == 0 )
				{
					MessageBox.Show("Debe Ingresar el Nombre del contacto  ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				
					this.uGridContacto.ActiveRow = dr;
					this.uGridContacto.ActiveRow.Selected = true;
					return;
				}

				if ( dr.Cells["Empresa"].Value.ToString().Trim().Length == 0)
				{
					MessageBox.Show("Debe Ingresar la Empresa del contacto ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				
					this.uGridContacto.ActiveRow = dr;
					this.uGridContacto.ActiveRow.Selected = true;
					return;
				}
				#region Telefonos Vacios
				if (dr.Cells["Convencional"].Value.ToString().Trim().Length == 0 && dr.Cells["Celular"].Value.ToString().Trim().Length == 0)
				{
					MessageBox.Show("Debe Ingresar el Telefono Convencional o el Telefono Celular del Contacto  ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				
					this.uGridContacto.ActiveRow = dr;
					this.uGridContacto.ActiveRow.Selected = true;
					return;
				}
				#endregion Telefonos Vacios

					if(dr.Cells["EMail"].Value.ToString().Length==0)
					{
						
						MessageBox.Show("Ingrese el email del contacto", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						
						this.uGridContacto.ActiveRow = dr;
						this.uGridContacto.ActiveRow.Selected = true;
						return;
					}
			}
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drValidaContactos in this.uGridContacto.Rows.All)
			{
				#region Telefono
				if (drValidaContactos.Cells["Convencional"].Value.ToString().Trim().Length > 0)
				{
					if (drValidaContactos.Cells["Convencional"].Value.ToString().Trim().Length != 9)
					{
						MessageBox.Show("El Telefono Convencional debe tener 9 Digitos.\n\nLos dos primeros corresponden al Codigo de Provincia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					
						this.uGridContacto.ActiveRow = drValidaContactos;
						this.uGridContacto.ActiveRow.Selected = true;
						return;
					}
				}
				#endregion Telefono
				#region Celular
				if (drValidaContactos.Cells["Celular"].Value.ToString().Trim().Length > 0)
				{
					if (drValidaContactos.Cells["Celular"].Value.ToString().Trim().Length != 10)
					{
						MessageBox.Show("El Telefono Celular debe tener 10 digitos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					
						this.uGridContacto.ActiveRow = drValidaContactos;
						this.uGridContacto.ActiveRow.Selected = true;
						return;
					}
				}
				#endregion Celular
				if(drValidaContactos.Cells["Email"].Value.ToString().Length==0)
				{
						
					MessageBox.Show("Ingrese el email del contacto", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridContacto.ActiveRow = drValidaContactos;
					this.uGridContacto.ActiveRow.Selected = true;
					return;
				}
				if (!Validacion.bEMail(drValidaContactos.Cells["Email"].Value.ToString())) 
				{
					this.uGridContacto.ActiveRow = drValidaContactos;
					this.uGridContacto.ActiveRow.Selected = true;
					return;
				}
			}
			#endregion validar campos datagrid 

			#region Guarda Registros contacto
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridContacto.Rows.All)
			{
				string sSQLDetalle = string.Format("Exec GuardaContacto {0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}'", 
				  (int)dr.Cells["idContacto"].Value,
					dr.Cells["Contacto"].Value.ToString(),
					dr.Cells["Empresa"].Value.ToString(),
					dr.Cells["Convencional"].Value.ToString(),
					dr.Cells["Extension"].Value.ToString(),
					dr.Cells["Celular"].Value.ToString(),
					dr.Cells["EMail"].Value.ToString(),
					dr.Cells["Notas"].Value.ToString()
					);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalle);
			}
			#endregion Guarda Registros contacto
			
			this.Close();	
		}
		private void frmContacto_Load(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Exec ConsultaGeneralContacto ");
			this.uGridContacto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);			

		}

		private void uGridContacto_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
//			this.uGridContacto.ActiveRow.Cells["idContacto"].Value.ToString()==0;
			
		}

		private void uGridContacto_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (uGridContacto.ActiveCell.Column.Key.ToString().ToUpper() == "CONVENCIONAL" || uGridContacto.ActiveCell.Column.Key.ToString().ToUpper() == "CELULAR" || uGridContacto.ActiveCell.Column.Key.ToString().ToUpper() == "EXTENSION")
			{
				Funcion.CamposNumericos(sender, e);}
		}

		private void uGridContacto_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender,e,this.uGridContacto);
		}

		private void uGridContacto_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
				e.Row.Cells["idContacto"].Value = 0;
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
					this.Close();
		}

		private void frmContacto_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();		
		}
	}
}
