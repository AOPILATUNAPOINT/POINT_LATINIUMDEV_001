using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using Infragistics.Shared;
using Infragistics.Win;
using Infragistics.Win.UltraWinGrid;
using System.Diagnostics;
using C1.Data;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmProductosDescuento.
	/// </summary>
	public class frmProductosDescuento : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsSeteoF;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.GroupBox gbBotones;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnEditar;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmProductosDescuento()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descuento");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total", 0);
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descuento");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmProductosDescuento));
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.gbBotones = new System.Windows.Forms.GroupBox();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
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
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource2;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 66;
			ultraGridColumn2.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 177;
			ultraGridColumn3.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 205;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance2;
			ultraGridColumn4.Format = "#,##0.00";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.PromptChar = ' ';
			ultraGridColumn4.Width = 69;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance3;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.NullText = "0.00";
			ultraGridColumn5.Width = 81;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance4;
			ultraGridColumn6.DataType = typeof(System.Decimal);
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 79;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			ultraGridBand1.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance5;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance6.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance6;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance7.FontData.BoldAsString = "True";
			appearance7.FontData.Name = "Arial";
			appearance7.FontData.SizeInPoints = 9F;
			appearance7.ForeColor = System.Drawing.Color.White;
			appearance7.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance7;
			appearance8.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance8.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 8);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(632, 272);
			this.ultraGrid1.TabIndex = 15;
			this.ultraGrid1.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_CellChange);
			this.ultraGrid1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ultraGrid1_KeyPress);
			this.ultraGrid1.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid1_BeforeRowsDeleted);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(System.Decimal);
			ultraDataColumn5.DataType = typeof(System.Decimal);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5});
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnNuevo.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 292);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 16;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// gbBotones
			// 
			this.gbBotones.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gbBotones.Location = new System.Drawing.Point(-216, 280);
			this.gbBotones.Name = "gbBotones";
			this.gbBotones.Size = new System.Drawing.Size(1088, 8);
			this.gbBotones.TabIndex = 264;
			this.gbBotones.TabStop = false;
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnCancelar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(156, 292);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 265;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEditar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnEditar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(80, 292);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 266;
			this.btnEditar.Text = "Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// frmProductosDescuento
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(648, 318);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.gbBotones);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.ultraGrid1);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmProductosDescuento";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Productos Descuento";
			this.Load += new System.EventHandler(this.frmProductosDescuento_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion


		#region Variables
		bool bNuevo = false;
		bool bEdicion = false;
		bool bEliminar = false;
		int iArticulo = 0;
		int ValiArticulo = 0;
		decimal dPrecio = 0.00m;
		decimal dDescuento = 0.00m;
		decimal dTotal = 0.00m;
		string sCodigoArt = "";
		string sArticulo = "";
		private Acceso miAcceso;
		#endregion Variables

		private void UnloadMe()
		{
			this.Close();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void frmProductosDescuento_Load(object sender, System.EventArgs e)
		{
			#region Seguridad
			miAcceso = new Acceso(cdsSeteoF, "1123");

			if (!Funcion.Permiso("1711", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Descuento Productos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;		
			}
			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Editar) this.btnEditar.Enabled = true;
			if (miAcceso.Eliminar) bEliminar = true;			
			#endregion Seguridad

			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Select idArticulo, Codigo, Articulo, Precio, Descuento, Total From ProductosDescuento"));

			FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);
			this.ultraGrid1.DisplayLayout.Bands[0].Columns["Descuento"].CellActivation = Activation.AllowEdit; // Seguridad Cambia Descuentos
			bNuevo = true;

			this.btnEditar.Enabled = false;
			this.btnNuevo.Enabled = false;
			this.btnCancelar.Enabled = true;
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);
			bNuevo = false;
			bEdicion = false;
			EncerarValores();
			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Editar) this.btnEditar.Enabled = true;
			this.btnCancelar.Enabled = false;
		}

		private void EncerarValores()
		{
			iArticulo = 0;
			ValiArticulo = 0;
			dPrecio = 0.00m;
			dDescuento = 0.00m;
			dTotal = 0.00m;
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (bNuevo || bEdicion)
			{
				if (e.KeyCode == Keys.F3)
				{
					string stArticulo = "";
					stArticulo = this.ultraGrid1.ActiveRow.Cells["Codigo"].Value.ToString();
					using (BuscaArticulos miBusqueda = new BuscaArticulos(sArticulo, 1, 1))
					{
						if (DialogResult.OK == miBusqueda.ShowDialog())
						{
							iArticulo = miBusqueda.idArticulo;
							sCodigoArt = miBusqueda.sCodigo;
							sArticulo = miBusqueda.sArticulo;
							PrecioArticulo(iArticulo);

							ultraGrid1.ActiveRow.Cells["idArticulo"].Value = iArticulo;
							ultraGrid1.ActiveRow.Cells["Codigo"].Value = sCodigoArt;
							ultraGrid1.ActiveRow.Cells["Articulo"].Value = sArticulo;
						}
					}		
				}
			}
		}

		private void PrecioArticulo(int idArticulo)
		{
			dPrecio = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec PrecioProductoDescuento {0}", idArticulo));
			ultraGrid1.ActiveRow.Cells["Precio"].Value = dPrecio;
		}

		private void ActualizaDatos(int idArticulo, string Codigo, string Articulo, decimal Precio)
		{
			dDescuento = Convert.ToDecimal(this.ultraGrid1.ActiveRow.Cells["Descuento"].Value);

			#region Validacion
			if (idArticulo > 0)
			{
				if (bNuevo)
				{
					string sSQL = string.Format("Select COUNT(1) From ProductosDescuento Where idArticulo = {0}", idArticulo);
					int ValidArticulo = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
					if (ValidArticulo == 1)
					{
						MessageBox.Show("El articulo ya Existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}
				}
			}

			if (idArticulo == 0 || Articulo == "")
			{
				MessageBox.Show("El Articulo no ha sido ingresado o no Existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			else if (Codigo == "")
			{
				MessageBox.Show("El Codigo del Articulo no ha sido ingresado o no Existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			else if (Convert.ToDecimal(Precio) <= 0)
			{
				MessageBox.Show("El Precio es Menor o igual a Cero", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			else if(Convert.ToDecimal(dDescuento) <= 0)
			{
				MessageBox.Show("El Descuento es Menor o igual a Cero", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			else if (Convert.ToDecimal(dDescuento) >= Convert.ToDecimal(Precio))
			{
				MessageBox.Show("El Descuento no puede ser Mayo o Igual al Precio", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
				#endregion Validacion
			else 
			{
				string sSQL = string.Format("Exec ActualizaProductosDescuentos {0}, '{1}', '{2}', {3}, {4}", idArticulo, Codigo, Articulo, Precio, dDescuento);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);

				this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Select idArticulo, Codigo, Articulo, Precio, Descuento, Total From ProductosDescuento"));
				EncerarValores();
			}
		}
		private void ActualizarValores()
		{
			ActualizaDatos(iArticulo, sCodigoArt, sArticulo, Convert.ToDecimal(dPrecio));
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bNuevo && e.Cell.Row.Cells["Descuento"].Value != System.DBNull.Value)
			{
				ActualizaDatos(iArticulo, sCodigoArt, sArticulo, Convert.ToDecimal(dPrecio));
				this.btnCancelar_Click(sender, e);
			}
			
			if (bEdicion)
			{
				ConsultaValores();
				ActualizaDatos(iArticulo, sCodigoArt, sArticulo, Convert.ToDecimal(dPrecio));
				this.btnCancelar_Click(sender, e);
			}
		}

		private void ultraGrid1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			//if (e.KeyChar == 13) this.ActualizarValores();
		}

		private void ConsultaValores()
		{
			ValiArticulo = (int)this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value;
			string sSQLRet = string.Format("Select idArticulo, Codigo, Articulo, Precio, Descuento, Total From ProductosDescuento Where idArticulo = {0}", ValiArticulo);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQLRet, true);
			dr.Read();
			if (dr.HasRows)
			{	
				iArticulo = dr.GetInt32(0);
				sCodigoArt  = dr.GetString(1);
				sArticulo = dr.GetString(2);
				dPrecio  = dr.GetDecimal(3);
				dDescuento = dr.GetDecimal(4);
				dTotal = dr.GetDecimal(5);
			}
			dr.Close();
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			bEdicion = true;
			FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);
			this.ultraGrid1.DisplayLayout.Bands[0].Columns["Descuento"].CellActivation = Activation.AllowEdit;

			this.btnNuevo.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnCancelar.Enabled = true;
		}

		private void ultraGrid1_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{

		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Descuento", 2);

			#region Formato numerico
			CultureInfo cultureENUS = CultureInfo.CreateSpecificCulture("ja-jp");
			string stMoneda = "#,##0.";
			string stMascara = "nnn,nnn,nnn.";
			string stCeros = "0000000000";
			string stNumero = "nnnnnnnnnn";
			string stFormato = "";
			string stInput = "";
			#endregion Formato numerico
			
			#region Decimales para el total
			stFormato = stMoneda + stCeros.Substring(0, 2); //2
			stInput = stMascara + stNumero.Substring(0, 2); //2			
			//this.txtTotal.MaskInput = stInput;
			#endregion Decimales para el total		
		}

		private void ultraGrid1_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if(bEliminar)
			{
				if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de borrar la línea seleccionada?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))				
				{
					e.DisplayPromptMsg = false;
					int DeliArticulo = (int)this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value;
					if (DeliArticulo > 0)
					{
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Delete From ProductosDescuento Where idArticulo = {0}", (int)this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value));
					}
					else
					{
						MessageBox.Show("Seleccione el producto a Eliminar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);						
						return;
					}
				}
			}
			else
			{
				MessageBox.Show("No tiene Acceso a Eliminar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.DisplayPromptMsg = false;
				return;
			}
		}
	}
}
