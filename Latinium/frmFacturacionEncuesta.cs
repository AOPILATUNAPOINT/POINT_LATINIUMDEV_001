using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmFacturacionEncuesta.
	/// </summary>
	public class frmFacturacionEncuesta : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.GroupBox groupBox1;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtOtros;
		private Infragistics.Win.UltraWinGrid.UltraGrid udgCompraEncuestaObtenerActivos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsCompraEncuestaObtenerActivos;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor optEncuesta; // Cambiado a tipo object
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmFacturacionEncuesta()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompraEncuesta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Seleccionar", 0);
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompraEncuesta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.txtOtros = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.udgCompraEncuestaObtenerActivos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsCompraEncuestaObtenerActivos = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.optEncuesta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			((System.ComponentModel.ISupportInitialize)(this.txtOtros)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udgCompraEncuestaObtenerActivos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsCompraEncuestaObtenerActivos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEncuesta)).BeginInit();
			this.SuspendLayout();
			// 
			// btnAceptar
			// 
			this.btnAceptar.Location = new System.Drawing.Point(8, 288);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(88, 23);
			this.btnAceptar.TabIndex = 2;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Location = new System.Drawing.Point(104, 288);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(88, 23);
			this.btnCancelar.TabIndex = 3;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.Visible = false;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// txtOtros
			// 
			this.txtOtros.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtOtros.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtOtros.Enabled = false;
			this.txtOtros.Location = new System.Drawing.Point(8, 280);
			this.txtOtros.MaxLength = 20;
			this.txtOtros.Name = "txtOtros";
			this.txtOtros.Size = new System.Drawing.Size(184, 21);
			this.txtOtros.TabIndex = 1;
			this.txtOtros.Visible = false;
			this.txtOtros.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtOtros_KeyDown);
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 264);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(202, 8);
			this.groupBox1.TabIndex = 141;
			this.groupBox1.TabStop = false;
			// 
			// udgCompraEncuestaObtenerActivos
			// 
			this.udgCompraEncuestaObtenerActivos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.udgCompraEncuestaObtenerActivos.Cursor = System.Windows.Forms.Cursors.Default;
			this.udgCompraEncuestaObtenerActivos.DataSource = this.udsCompraEncuestaObtenerActivos;
			appearance1.BackColor = System.Drawing.Color.White;
			this.udgCompraEncuestaObtenerActivos.DisplayLayout.Appearance = appearance1;
			this.udgCompraEncuestaObtenerActivos.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 86;
			ultraGridColumn2.AutoEdit = false;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn2.Header.Caption = "Descripción";
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Width = 135;
			ultraGridColumn3.DataType = typeof(bool);
			ultraGridColumn3.DefaultCellValue = false;
			ultraGridColumn3.Header.Caption = "...";
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn3.Width = 28;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			ultraGridBand1.Header.Caption = "Detalle";
			this.udgCompraEncuestaObtenerActivos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.udgCompraEncuestaObtenerActivos.DisplayLayout.GroupByBox.Hidden = true;
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.udgCompraEncuestaObtenerActivos.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.udgCompraEncuestaObtenerActivos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.udgCompraEncuestaObtenerActivos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.udgCompraEncuestaObtenerActivos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.udgCompraEncuestaObtenerActivos.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 10F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.udgCompraEncuestaObtenerActivos.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.udgCompraEncuestaObtenerActivos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance5.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance5.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgCompraEncuestaObtenerActivos.DisplayLayout.Override.RowAlternateAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgCompraEncuestaObtenerActivos.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgCompraEncuestaObtenerActivos.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.udgCompraEncuestaObtenerActivos.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.udgCompraEncuestaObtenerActivos.FlatMode = true;
			this.udgCompraEncuestaObtenerActivos.Font = new System.Drawing.Font("Tahoma", 7.75F);
			this.udgCompraEncuestaObtenerActivos.Location = new System.Drawing.Point(8, 8);
			this.udgCompraEncuestaObtenerActivos.Name = "udgCompraEncuestaObtenerActivos";
			this.udgCompraEncuestaObtenerActivos.Size = new System.Drawing.Size(184, 248);
			this.udgCompraEncuestaObtenerActivos.TabIndex = 763;
			this.udgCompraEncuestaObtenerActivos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.udgCompraEncuestaObtenerActivos_AfterCellUpdate);
			this.udgCompraEncuestaObtenerActivos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.udgCompraEncuestaObtenerActivos_InitializeLayout);
			// 
			// udsCompraEncuestaObtenerActivos
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.udsCompraEncuestaObtenerActivos.Band.Columns.AddRange(new object[] {
																																								ultraDataColumn1,
																																								ultraDataColumn2});
			// 
			// optEncuesta
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEncuesta.Appearance = appearance8;
			this.optEncuesta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.optEncuesta.Enabled = false;
			this.optEncuesta.FormatString = "";
			this.optEncuesta.Location = new System.Drawing.Point(69, 150);
			this.optEncuesta.MaskInput = "nn";
			this.optEncuesta.Name = "optEncuesta";
			this.optEncuesta.PromptChar = ' ';
			this.optEncuesta.Size = new System.Drawing.Size(64, 21);
			this.optEncuesta.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.optEncuesta.TabIndex = 765;
			this.optEncuesta.Visible = false;
			// 
			// frmFacturacionEncuesta
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(202, 320);
			this.ControlBox = false;
			this.Controls.Add(this.optEncuesta);
			this.Controls.Add(this.udgCompraEncuestaObtenerActivos);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.txtOtros);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.Name = "frmFacturacionEncuesta";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Encuesta";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmFacturacionEncuesta_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmFacturacionEncuesta_Closing);
			this.Load += new System.EventHandler(this.frmFacturacionEncuesta_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtOtros)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udgCompraEncuestaObtenerActivos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsCompraEncuestaObtenerActivos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEncuesta)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmFacturacionEncuesta_Load(object sender, System.EventArgs e)
		{
			CompraEncuestaObtenerActivos();
		}

		private void CompraEncuestaObtenerActivos()
		{
			string squery = "EXEC CompraEncuestaObtenerActivos";
			FuncionesProcedimientos.dsGeneralModulo(squery, this.udgCompraEncuestaObtenerActivos);
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void frmFacturacionEncuesta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}


		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			// Verificar si el UltraGrid está vacío
			if (this.udgCompraEncuestaObtenerActivos.Rows.Count == 0)
			{
				MessageBox.Show("La tabla está vacía.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}

			// Contar cuántos checks están seleccionados
			int selectedCount = 0;
			Infragistics.Win.UltraWinGrid.UltraGridRow selectedRow = null; // Variable para almacenar la fila seleccionada

			foreach (Infragistics.Win.UltraWinGrid.UltraGridRow row in this.udgCompraEncuestaObtenerActivos.Rows)
			{
				// Suponiendo que el campo "Seleccionar" es de tipo bool
				if (Convert.ToBoolean(row.Cells["Seleccionar"].Value))
				{
					selectedCount++;
					selectedRow = row; // Guardar la fila seleccionada
				}
			}

			// Validar que solo uno esté seleccionado
			if (selectedCount != 1)
			{
				MessageBox.Show("Debe seleccionar exactamente un elemento.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}

			// Asignar el valor a la variable optEncuesta
			optEncuesta.Value = (int)selectedRow.Cells["idCompraEncuesta"].Value; // O el campo que corresponda

			// Si todo está bien, establece el resultado del diálogo
			this.DialogResult = DialogResult.OK;
		}


		private void optEncuesta_ValueChanged(object sender, System.EventArgs e)
		{
		}

		private void txtOtros_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.btnAceptar.Focus();
		}

		private void frmFacturacionEncuesta_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
		}

		protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
		{
			// Verificar si la tecla presionada es la tecla Escape
			if (keyData == Keys.Escape)
			{
				// Devolver true para indicar que el evento de tecla ha sido manejado
				return true;
			}

			// Verificar si las teclas presionadas son Alt + F4
			if (keyData == (Keys.Alt | Keys.F4))
			{
				// Devolver true para indicar que el evento de tecla ha sido manejado
				return true;
			}

			// Llamar al método de la clase base para otras teclas
			return base.ProcessCmdKey(ref msg, keyData);
		}

		private void udgCompraEncuestaObtenerActivos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			// Verificar si la celda actual es la de "Seleccionar"
			if (e.Cell.Column.Key == "Seleccionar")
			{
				// Contar cuántos checks están seleccionados
				int selectedCount = 0;

				foreach (Infragistics.Win.UltraWinGrid.UltraGridRow row in this.udgCompraEncuestaObtenerActivos.Rows)
				{
					// Suponiendo que el campo "Seleccionar" es de tipo bool
					if (Convert.ToBoolean(row.Cells["Seleccionar"].Value))
					{
						selectedCount++;
					}
				}

				// Si se selecciona más de uno, mostrar mensaje y deshacer el cambio
				if (selectedCount > 1)
				{
					MessageBox.Show("Solo se puede seleccionar un ítem.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);

					// Desmarcar la celda actual
					e.Cell.Value = false; // Desmarcar la celda

					// Si deseas forzar que el grid se vuelva a dibujar
					this.udgCompraEncuestaObtenerActivos.Refresh(); // Refrescar el grid
				}
			}
		}

		private void udgCompraEncuestaObtenerActivos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
