using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmFacturacionFaltantes.
	/// </summary>
	public class frmFacturacionFaltantes : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnAceptar;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridPersonalConteo;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValorTotal;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcentaje;
		private System.Windows.Forms.Label label55;
		private System.Windows.Forms.Label label1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmFacturacionFaltantes()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}
		int idCierre = 0;
		DateTime Fecha ;
		int iBodega;
		int iEstado ;
		decimal totalFaltantes;
		public frmFacturacionFaltantes(int IdCierre, DateTime dtFecha, int Bodega, decimal DFaltantes, int IEstado)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			idCierre = IdCierre;			
			Fecha = dtFecha;
			iBodega = Bodega;			
			totalFaltantes = DFaltantes;
			iEstado = IEstado;
			//
			// TODO: Add any constructor code after InitializeComponent call
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PersonalCierreConteoInventario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuotas");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorCuotas");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Anticipo");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PersonalCierreConteoInventario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuotas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorCuotas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Anticipo");
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.uGridPersonalConteo = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtValorTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPorcentaje = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label55 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.uGridPersonalConteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcentaje)).BeginInit();
			this.SuspendLayout();
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(104, 376);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(76, 23);
			this.btnCancelar.TabIndex = 409;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnAceptar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAceptar.Location = new System.Drawing.Point(16, 376);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(76, 23);
			this.btnAceptar.TabIndex = 408;
			this.btnAceptar.Text = "Aceptar";
			this.btnAceptar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// uGridPersonalConteo
			// 
			this.uGridPersonalConteo.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridPersonalConteo.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridPersonalConteo.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridPersonalConteo.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 79;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 186;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn5.CellAppearance = appearance2;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 68;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance3;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 70;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn7.CellAppearance = appearance4;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 53;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance5;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 85;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 54;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			ultraGridBand1.Header.Caption = "Plazos";
			this.uGridPersonalConteo.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance6.ForeColor = System.Drawing.Color.Black;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPersonalConteo.DisplayLayout.Override.ActiveRowAppearance = appearance6;
			this.uGridPersonalConteo.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridPersonalConteo.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance7.BackColor = System.Drawing.Color.Transparent;
			this.uGridPersonalConteo.DisplayLayout.Override.CardAreaAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 8F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridPersonalConteo.DisplayLayout.Override.HeaderAppearance = appearance8;
			this.uGridPersonalConteo.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance9.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance9.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPersonalConteo.DisplayLayout.Override.RowAlternateAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPersonalConteo.DisplayLayout.Override.RowSelectorAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPersonalConteo.DisplayLayout.Override.SelectedRowAppearance = appearance11;
			this.uGridPersonalConteo.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridPersonalConteo.Location = new System.Drawing.Point(0, 48);
			this.uGridPersonalConteo.Name = "uGridPersonalConteo";
			this.uGridPersonalConteo.Size = new System.Drawing.Size(616, 320);
			this.uGridPersonalConteo.TabIndex = 407;
			this.uGridPersonalConteo.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridPersonalConteo_CellChange);
			this.uGridPersonalConteo.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridPersonalConteo_AfterCellUpdate);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(System.Decimal);
			ultraDataColumn6.DataType = typeof(System.Decimal);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn9.DataType = typeof(bool);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9});
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
			// txtValorTotal
			// 
			this.txtValorTotal.Anchor = System.Windows.Forms.AnchorStyles.Top;
			this.txtValorTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValorTotal.Enabled = false;
			this.txtValorTotal.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtValorTotal.Location = new System.Drawing.Point(232, 16);
			this.txtValorTotal.Name = "txtValorTotal";
			this.txtValorTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValorTotal.PromptChar = ' ';
			this.txtValorTotal.Size = new System.Drawing.Size(64, 22);
			this.txtValorTotal.TabIndex = 411;
			// 
			// txtPorcentaje
			// 
			this.txtPorcentaje.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtPorcentaje.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcentaje.Enabled = false;
			this.txtPorcentaje.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtPorcentaje.Location = new System.Drawing.Point(80, 16);
			this.txtPorcentaje.Name = "txtPorcentaje";
			this.txtPorcentaje.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcentaje.PromptChar = ' ';
			this.txtPorcentaje.Size = new System.Drawing.Size(64, 22);
			this.txtPorcentaje.TabIndex = 410;
			// 
			// label55
			// 
			this.label55.AutoSize = true;
			this.label55.BackColor = System.Drawing.Color.Transparent;
			this.label55.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label55.Location = new System.Drawing.Point(16, 19);
			this.label55.Name = "label55";
			this.label55.Size = new System.Drawing.Size(57, 17);
			this.label55.TabIndex = 412;
			this.label55.Text = "Porcentaje";
			this.label55.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(168, 19);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(58, 17);
			this.label1.TabIndex = 413;
			this.label1.Text = "Valor Total";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// frmFacturacionFaltantes
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(616, 406);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label55);
			this.Controls.Add(this.txtValorTotal);
			this.Controls.Add(this.txtPorcentaje);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.uGridPersonalConteo);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmFacturacionFaltantes";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Facturacion Faltantes";
			this.Load += new System.EventHandler(this.frmFacturacionFaltantes_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridPersonalConteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcentaje)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private bool _cargando = false;
		private void frmFacturacionFaltantes_Load(object sender, System.EventArgs e)
		{
			_cargando = true;
			if (iEstado == 1) 
			{
				FuncionesProcedimientos.EstadoGrids(true, this.uGridPersonalConteo);
				
			}
			else
			{
				FuncionesProcedimientos.EstadoGrids(false, this.uGridPersonalConteo);
				this.btnAceptar.Enabled = false ;
			}

			string sSQL = string.Format("Exec ListaPersonalCuotas {0},{1}, '{2}'", idCierre,iBodega, Fecha.ToString("yyyyMMdd"));
			this.uGridPersonalConteo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);	

			int totalPersonas = uGridPersonalConteo.Rows.Count;

			this.txtValorTotal.Text = totalFaltantes.ToString("0.00");
			this.txtPorcentaje.Text = "100.00";

			if (totalPersonas == 0 || totalFaltantes == 0)
			{
				_cargando = false;
				return;
			}

			decimal porcentajePorPersona = Math.Round(100m / totalPersonas, 3);

			foreach (UltraGridRow row in uGridPersonalConteo.Rows)
			{
				row.Cells["Porcentaje"].Value = porcentajePorPersona;

				decimal valor = Math.Round(totalFaltantes * (porcentajePorPersona / 100), 2);
				row.Cells["Valor"].Value = valor;
				row.Cells["Cuotas"].Value = 1;
				row.Cells["ValorCuotas"].Value = valor; 
			}
			_cargando = false;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);			
		}

		private void uGridPersonalConteo_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (_cargando) return;
			if (e.Cell.Column.Key == "Porcentaje" || e.Cell.Column.Key == "Cuotas")
			{
				ActualizarTotalesEnGrid();
			}
		}
		private void ActualizarTotalesEnGrid()
		{
			decimal sumaPorcentaje = 0;
			decimal sumaValor = 0;
			bool excedePorcentaje = false;

			for (int i = 0; i < uGridPersonalConteo.Rows.Count; i++)
			{
				UltraGridRow row = uGridPersonalConteo.Rows[i];
				decimal porcentaje = 0;
				int cuotas = 1;

				try
				{
					if (row.Cells["Porcentaje"].Value != null && row.Cells["Porcentaje"].Value.ToString() != "")
						porcentaje = Convert.ToDecimal(row.Cells["Porcentaje"].Value);
				}
				catch
				{
					porcentaje = 0;
				}

				try
				{
					if (row.Cells["Cuotas"].Value != null && row.Cells["Cuotas"].Value.ToString() != "")
						cuotas = Convert.ToInt32(row.Cells["Cuotas"].Value);
				}
				catch
				{
					cuotas = 1;
				}

				if (cuotas <= 0) cuotas = 1;

				// Validación de porcentaje acumulado
				if ((sumaPorcentaje + porcentaje) > 100m)
				{
					decimal permitido = 100m - sumaPorcentaje;

					row.Cells["Porcentaje"].Value = permitido;  // Ajusta al máximo permitido
					porcentaje = permitido;
					excedePorcentaje = true;
				}

				sumaPorcentaje += porcentaje;

				decimal valor = Math.Round(totalFaltantes * (porcentaje / 100), 2);
				row.Cells["Valor"].Value = valor;
				row.Cells["ValorCuotas"].Value = Math.Round(valor / cuotas, 2);

				sumaValor += valor;
			}

			// Mostrar advertencia si excedía
			if (excedePorcentaje)
			{
				MessageBox.Show("El porcentaje total no puede superar el 100%. Se ha ajustado automáticamente el último valor ingresado.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
			}

			// Actualizar totales visuales
			txtPorcentaje.Text = sumaPorcentaje.ToString("0.00");
			txtValorTotal.Text = sumaValor.ToString("0.00");

			if (Math.Round(sumaPorcentaje, 2) == 100m)
				txtPorcentaje.BackColor = Color.LightGreen;
			else
				txtPorcentaje.BackColor = Color.LightCoral;
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			if (uGridPersonalConteo.Rows.Count == 0)
			{
				MessageBox.Show("No hay datos para guardar.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}
			decimal porcentajeTotal = 0;
			foreach (UltraGridRow row in uGridPersonalConteo.Rows)
			{
				if (row.Cells["Porcentaje"].Value != DBNull.Value)
				{
					porcentajeTotal += Convert.ToDecimal(row.Cells["Porcentaje"].Value);
				}
			}

			if (Math.Round(porcentajeTotal, 2) != 100)
			{
				MessageBox.Show("La suma de los porcentajes debe ser exactamente 100%.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}
			foreach (UltraGridRow row in uGridPersonalConteo.Rows)
			{
				if (row.Cells["PersonalCierreConteoInventario"].Value == DBNull.Value ||
					row.Cells["Porcentaje"].Value == DBNull.Value ||row.Cells["Valor"].Value == DBNull.Value ||
					row.Cells["Cuotas"].Value == DBNull.Value ||row.Cells["ValorCuotas"].Value == DBNull.Value)
				{
					MessageBox.Show("Todos los campos deben estar completos para guardar.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					uGridPersonalConteo.ActiveRow = row;
					uGridPersonalConteo.ActiveRow.Selected = true;
					return;
				}
				string sSQL = string.Format("Exec ActualizaCuotasFaltantes {0}, {1}, {2}, {3}, {4}",
					(int)row.Cells["PersonalCierreConteoInventario"].Value,(decimal)row.Cells["Porcentaje"].Value,
					(decimal)row.Cells["Valor"].Value,(int)row.Cells["Cuotas"].Value,(decimal)row.Cells["ValorCuotas"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);
			}
			string sSQL3 = string.Format("Exec GenerarPagosFaltantes {0}",idCierre);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL3);


			MessageBox.Show("Datos guardados correctamente.", "Confirmación", MessageBoxButtons.OK, MessageBoxIcon.Information);
			this.DialogResult = DialogResult.OK;
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void uGridPersonalConteo_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			try
			{
				if (e.Cell.Column.Key == "Anticipo")
				{
					this.uGridPersonalConteo.UpdateData();

					UltraGridRow row = e.Cell.Row;

					if (row.Cells["PersonalCierreConteoInventario"].Value != DBNull.Value)
					{
						int idPersonal = (int)row.Cells["PersonalCierreConteoInventario"].Value;
						bool firmo = Convert.ToBoolean(row.Cells["Anticipo"].Value); 
						int valorFirmo = firmo ? 1 : 0;

						string sSQL = string.Format(
							"UPDATE PersonalCierreConteoInventario SET FirmoAnticipo = {0} WHERE PersonalCierreConteoInventario = {1}",
							valorFirmo, idPersonal
							);

						Funcion.EjecutaSQL(cdsSeteoF, sSQL);
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error al actualizar FirmoAnticipo: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}
	}
}
