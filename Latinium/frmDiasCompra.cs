using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmDiasCompra.
	/// </summary>
	public class frmDiasCompra : System.Windows.Forms.Form
	{
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnAceptar;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridDetalle;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		int idCompra = 0;
		bool  bHabita = true;
		int iTipo=1;

		public frmDiasCompra(int IdCompra ,int ITipo)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			idCompra =IdCompra;
			iTipo=ITipo;

			

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		public frmDiasCompra()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("pagos");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dias");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("pagos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("dias");
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.uGridDetalle = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 288);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(280, 8);
			this.groupBox1.TabIndex = 450;
			this.groupBox1.TabStop = false;
			// 
			// btnCancelar
			// 
			this.btnCancelar.Location = new System.Drawing.Point(80, 304);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(75, 24);
			this.btnCancelar.TabIndex = 449;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnAceptar
			// 
			this.btnAceptar.Location = new System.Drawing.Point(0, 304);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(75, 24);
			this.btnAceptar.TabIndex = 448;
			this.btnAceptar.Text = "Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// uGridDetalle
			// 
			this.uGridDetalle.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDetalle.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDetalle.DataSource = this.ultraDataSource4;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridDetalle.DisplayLayout.Appearance = appearance1;
			this.uGridDetalle.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 64;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn2.CellAppearance = appearance2;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn2.FormulaErrorAppearance = appearance3;
			ultraGridColumn2.Header.Caption = "Numero de Pagos";
			ultraGridColumn2.Header.VisiblePosition = 1;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn2.MaskLiteralsAppearance = appearance4;
			ultraGridColumn2.Width = 120;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn3.CellAppearance = appearance5;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn3.CellButtonAppearance = appearance6;
			ultraGridColumn3.Format = "";
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn3.FormulaErrorAppearance = appearance7;
			ultraGridColumn3.Header.Caption = "Días de Pago";
			ultraGridColumn3.Header.VisiblePosition = 2;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn3.MaskLiteralsAppearance = appearance8;
			ultraGridColumn3.Width = 120;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			ultraGridBand1.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			ultraGridBand1.SummaryFooterCaption = "Totales";
			this.uGridDetalle.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance9.ForeColor = System.Drawing.Color.Black;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalle.DisplayLayout.Override.ActiveRowAppearance = appearance9;
			this.uGridDetalle.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridDetalle.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridDetalle.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance10.BackColor = System.Drawing.Color.Transparent;
			this.uGridDetalle.DisplayLayout.Override.CardAreaAppearance = appearance10;
			this.uGridDetalle.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance11.FontData.BoldAsString = "True";
			appearance11.FontData.Name = "Arial";
			appearance11.FontData.SizeInPoints = 8.5F;
			appearance11.ForeColor = System.Drawing.Color.White;
			appearance11.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDetalle.DisplayLayout.Override.HeaderAppearance = appearance11;
			appearance12.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance12.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowAlternateAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowSelectorAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.SelectedRowAppearance = appearance14;
			this.uGridDetalle.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDetalle.Location = new System.Drawing.Point(8, 0);
			this.uGridDetalle.Name = "uGridDetalle";
			this.uGridDetalle.Size = new System.Drawing.Size(261, 280);
			this.uGridDetalle.TabIndex = 447;
			this.uGridDetalle.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridDetalle_AfterRowInsert);
			this.uGridDetalle.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridDetalle_BeforeRowsDeleted);
			this.uGridDetalle.BeforeRowInsert += new Infragistics.Win.UltraWinGrid.BeforeRowInsertEventHandler(this.uGridDetalle_BeforeRowInsert);
			this.uGridDetalle.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridDetalle_InitializeLayout);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn1.DataType = typeof(short);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3});
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
			// frmDiasCompra
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(277, 334);
			this.ControlBox = false;
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.uGridDetalle);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmDiasCompra";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "Días Crédito";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmDiasCompra_KeyDown);
			this.Load += new System.EventHandler(this.frmDiasCompra_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
				this.DialogResult = DialogResult.No;
		}

		private void frmDiasCompra_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void uGridDetalle_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Dias", 0);
			FuncionesProcedimientos.FormatoGrid(e, "Pagos", 0);
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			#region Validacion


			if (this.uGridDetalle.Rows.Count == 0)
			{
				MessageBox.Show("Ingrese el numero de Pagos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.uGridDetalle.Focus();			
			}
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGridDetalle.Rows.All)
			{
				if (dr.Cells["dias"].Value == System.DBNull.Value  )
				{
					MessageBox.Show("Ingrese Valor ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				if ((int)dr.Cells["dias"].Value == 0 )
				{
					MessageBox.Show("Ingrese Valor no puede  ser  cero  ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}	
				if (dr.Cells["pagos"].Value == System.DBNull.Value  )
				{
					MessageBox.Show("Ingrese Pagos ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				if ((int)dr.Cells["pagos"].Value == 0 )
				{
					MessageBox.Show("Ingrese Pagos no puede  ser  cero  ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}	
			}
			#endregion Validacion


			DialogResult = DialogResult.OK;
		}

		private void frmDiasCompra_Load(object sender, System.EventArgs e)
		{
			this.uGridDetalle.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaDiasDePagoCompras {0}", idCompra));

			if (iTipo==2){
				FuncionesProcedimientos.EstadoGrids(false, this.uGridDetalle);
				this.btnAceptar.Enabled=false;	
			} 

			if (iTipo==1)
			{
				if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec CXPValidaEstadoAutorizacion {0}", idCompra)) > 0)
				{
					bHabita = false;
					FuncionesProcedimientos.EstadoGrids(false, this.uGridDetalle);
				}
			}
		
		}

		private void uGridDetalle_BeforeRowInsert(object sender, Infragistics.Win.UltraWinGrid.BeforeRowInsertEventArgs e)
		{
			//int iTotalLineas = ultraGrid1.Rows.Count;

		}

		private void uGridDetalle_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			int iTotalLineas = 0 ;

			iTotalLineas=uGridDetalle.Rows.Count;

			e.Row.Cells["pagos"].Value = iTotalLineas;
			e.Row.Cells["idCompra"].Value = 0;
			e.Row.Cells["dias"].Value = 0;
		}

		private void uGridDetalle_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			#region Elimina
			for (int i = 0; i < e.Rows.Length; i++)
			{
				if ((int)e.Rows[i].Cells["idCompra"].Value > 0)
				{
					string sSQLValidaAu = string.Format("Exec ValidaAutorizacionDePagos {0}", (int)e.Rows[i].Cells["idCompra"].Value);
					int  iPagos = Funcion.iEscalarSQL(cdsSeteoF, sSQLValidaAu, true);
					if (iPagos > 0)
					{
						MessageBox.Show("No puede eliminar filas exiten pagos generados", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						e.Cancel = true;
						return;
					}
				}
			}
			e.DisplayPromptMsg = false; 
			#endregion Elimina
		}
	}
}
