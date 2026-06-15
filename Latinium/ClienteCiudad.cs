using System;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de ClienteCiudad.
	/// </summary>
	public class ClienteCiudad : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsCiudad;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private C1.Data.C1DataView cdvCiudad;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public ClienteCiudad()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: Agregar código de constructor después de llamar a InitializeComponent
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Método necesario para admitir el Diseñador, no se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ClienteCiudad", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ciudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			this.cdsCiudad = new C1.Data.C1DataSet();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.cdvCiudad = new C1.Data.C1DataView();
			((System.ComponentModel.ISupportInitialize)(this.cdsCiudad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCiudad)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsCiudad
			// 
			this.cdsCiudad.BindingContextCtrl = this;
			this.cdsCiudad.DataLibrary = "LibFacturacion";
			this.cdsCiudad.DataLibraryUrl = "";
			this.cdsCiudad.DataSetDef = "dsClienteTabla";
			this.cdsCiudad.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsCiudad.Name = "cdsCiudad";
			this.cdsCiudad.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCiudad.SchemaDef = null;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "ClienteCiudad";
			this.ultraGrid1.DataSource = this.cdsCiudad;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn2.Width = 172;
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn3.Width = 71;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(16, 16);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(264, 216);
			this.ultraGrid1.TabIndex = 0;
			this.ultraGrid1.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid1_BeforeRowsDeleted);
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Location = new System.Drawing.Point(32, 248);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.TabIndex = 1;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btCancelar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCancelar.Location = new System.Drawing.Point(184, 248);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.TabIndex = 2;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// cdvCiudad
			// 
			this.cdvCiudad.BindingContextCtrl = this;
			this.cdvCiudad.DataSet = this.cdsCiudad;
			this.cdvCiudad.TableName = "";
			this.cdvCiudad.TableViewName = "ClienteCiudad";
			// 
			// ClienteCiudad
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(296, 293);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.ultraGrid1);
			this.Name = "ClienteCiudad";
			this.Text = "Ciudades";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.ClienteCiudad_Closing);
			this.Load += new System.EventHandler(this.ClienteCiudad_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsCiudad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCiudad)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			try
			{
				this.cdsCiudad.Update();
			}
			catch (System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al Grabar");
				return;
			}
			#region Crea Cuidad en Empresas Relacionadas
			if (Funcion.bEjecutaSQL(cdsCiudad, "Exec SeteoGExiste 'FCLAER'"))
			{
				string stMensaje = Funcion.sEscalarSQL(cdsCiudad, "Exec GIM_ClienteCiudad "
					+"'"+MenuLatinium.stDirFacturacion+"'");

				if (stMensaje.Length > 0)
				{
					MessageBox.Show(stMensaje, "Información", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}
						
			}
			#endregion
			this.Close();
	

		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void VerificaBorrar(string IdCiudad, string stCiudad)
		{
			// Verificacion de Existencia en compras
			string stSelect = "SELECT COUNT(*) FROM Cliente WHERE Cliente.idCiudad = "
				+ IdCiudad;
			this.cdsCiudad.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = this.cdsCiudad.Schema.Connections[0];
			SqlCommand cmdCodigo = new SqlCommand (
				stSelect, (SqlConnection) miConeccion.DbConnection);
			int iCuenta = (int) cmdCodigo.ExecuteScalar();
			if (iCuenta > 0) 
			{
				MessageBox.Show("Tiene " + iCuenta.ToString().Trim() + " Clientes con la ciudad "
					+ stCiudad + ". Eliminación CANCELADA",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.cdsCiudad.Schema.Connections[0].Close();
				return;
			}
			this.cdsCiudad.Schema.Connections[0].Close();
			C1.Data.C1DataRow dr = this.cdsCiudad.TableViews["ClienteCiudad"].Rows.Find(int.Parse(IdCiudad));
			dr.Delete();
			// Fin de verificacion de Existencia de cliente en compras
		}

		private void ultraGrid1_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			e.DisplayPromptMsg = false;
			e.Cancel = true;
			string stCiudad = " Ciudades?";
			if (e.Rows.Length == 1) stCiudad = " Ciudad?";
			if(MessageBox.Show("¿Desea Borrar " + e.Rows.Length.ToString().Trim() 
				+ stCiudad, "Confirmación",
				MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
			{
				return;
			}

			Infragistics.Win.UltraWinGrid.UltraGridRow[] drCiudad = e.Rows;
			string ICiudad;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in drCiudad)
			{
				ICiudad = dr.Cells["idCiudad"].Text.Trim();
				stCiudad = dr.Cells["Ciudad"].Text.Trim();
				this.VerificaBorrar(ICiudad, stCiudad);
			}
		}

		private void ClienteCiudad_Load(object sender, System.EventArgs e)
		{
			Cursor = Cursors.Default;
			string stAudita = "Exec AuditaCrear 53, 6, 'Ciudad'";
			Funcion.EjecutaSQL(cdsCiudad, stAudita, true);
		}

		private void ClienteCiudad_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Ciudad'";
			Funcion.EjecutaSQL(cdsCiudad, stAudita, true);
		}
	}
}
