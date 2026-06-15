using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Casos.
	/// </summary>
	public class Casos : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.Windows.Forms.Label label1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private C1.Data.C1DataSet cdsCaso;
		private C1.Data.C1DataSet cdsTipoCaso;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private C1.Data.C1DataView cdvCaso;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoCaso;
		private C1.Data.C1DataSet cdsSeteoF;
    int VCaso;
		private System.Windows.Forms.Label label4;
		int compra;
		public Casos(int idCaso,int IdCompra)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			VCaso=idCaso;
			compra=IdCompra;
			
			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}
		public Casos()
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsCaso = new C1.Data.C1DataSet();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.label1 = new System.Windows.Forms.Label();
			this.cdsTipoCaso = new C1.Data.C1DataSet();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.cdvCaso = new C1.Data.C1DataView();
			this.cmbTipoCaso = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label4 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCaso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipoCaso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCaso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoCaso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
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
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(16, 136);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(232, 88);
			this.ultraGrid1.TabIndex = 2;
			this.ultraGrid1.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid1_BeforeRowsDeleted);
			// 
			// cdsCaso
			// 
			this.cdsCaso.BindingContextCtrl = this;
			this.cdsCaso.DataLibrary = "LibFacturacion";
			this.cdsCaso.DataLibraryUrl = "";
			this.cdsCaso.DataSetDef = "dsCaso";
			this.cdsCaso.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsCaso.Name = "cdsCaso";
			this.cdsCaso.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCaso.SchemaDef = null;
			this.cdsCaso.BeforeFill += new C1.Data.FillEventHandler(this.cdsCaso_BeforeFill);
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btCancelar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCancelar.Location = new System.Drawing.Point(152, 240);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.TabIndex = 8;
			this.btCancelar.Text = "&Cerrar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Location = new System.Drawing.Point(40, 240);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.TabIndex = 7;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(96, 40);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(88, 16);
			this.label1.TabIndex = 9;
			// 
			// cdsTipoCaso
			// 
			this.cdsTipoCaso.BindingContextCtrl = this;
			this.cdsTipoCaso.DataLibrary = "LibFacturacion";
			this.cdsTipoCaso.DataLibraryUrl = "";
			this.cdsTipoCaso.DataSetDef = "dsTipoCaso";
			this.cdsTipoCaso.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsTipoCaso.Name = "cdsTipoCaso";
			this.cdsTipoCaso.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsTipoCaso.SchemaDef = null;
			// 
			// label2
			// 
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(24, 40);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(56, 16);
			this.label2.TabIndex = 121;
			this.label2.Text = "# Caso:";
			// 
			// label3
			// 
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.Location = new System.Drawing.Point(40, 0);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(168, 16);
			this.label3.TabIndex = 122;
			this.label3.Text = "Asignacion a Tipos de Casos";
			// 
			// ultraButton1
			// 
			this.ultraButton1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.ultraButton1.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraButton1.Location = new System.Drawing.Point(16, 72);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.TabIndex = 123;
			this.ultraButton1.Text = "Tipo Caso";
			this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			this.errorProvider.DataMember = "";
			// 
			// cdvCaso
			// 
			this.cdvCaso.BindingContextCtrl = this;
			this.cdvCaso.DataSet = this.cdsCaso;
			this.cdvCaso.TableName = "";
			this.cdvCaso.TableViewName = "P_Caso";
			// 
			// cmbTipoCaso
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoCaso.Appearance = appearance6;
			this.cmbTipoCaso.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoCaso.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoCaso.DisplayMember = "TipoCaso";
			this.cmbTipoCaso.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoCaso.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoCaso.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoCaso.Location = new System.Drawing.Point(96, 72);
			this.cmbTipoCaso.Name = "cmbTipoCaso";
			this.cmbTipoCaso.Size = new System.Drawing.Size(144, 21);
			this.cmbTipoCaso.TabIndex = 148;
			this.cmbTipoCaso.ValueMember = "idTipoCaso";
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
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(200, 40);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(48, 16);
			this.label4.TabIndex = 149;
			// 
			// Casos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(264, 273);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.cmbTipoCaso);
			this.Controls.Add(this.ultraButton1);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.ultraGrid1);
			this.Name = "Casos";
			this.Text = "Casos";
			this.Load += new System.EventHandler(this.Casos_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCaso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipoCaso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCaso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoCaso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void Casos_Load(object sender, System.EventArgs e)
		{
				label1.Text=VCaso.ToString();
				label4.Text=compra.ToString();
				string stSelect="Select tp.TipoCaso,c.idCaso,c.Caso,c.idtipoCaso from P_Caso c,P_TipoCaso tp  where c.idtipocaso=tp.idtipocaso  and c.Caso="+VCaso+" and c.idcompra="+compra;
				ultraGrid1.DataSource=Funcion.dvProcedimiento(cdsSeteoF,stSelect);
				ultraGrid1.Rows.Band.Columns["idCaso"].Hidden=true;
				ultraGrid1.Rows.Band.Columns["Caso"].Hidden=true;
				ultraGrid1.Rows.Band.Columns["idTipoCaso"].Hidden=true;
				string stcombo="Select * from P_TipoCaso";
				cmbTipoCaso.DataSource=Funcion.dvProcedimiento(cdsSeteoF,stcombo);
				cmbTipoCaso.DisplayMember = "TipoCaso";
				cmbTipoCaso.ValueMember = "idTipoCaso";
				cmbTipoCaso.Rows.Band.Columns["idTipoCaso"].Hidden=true;
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
      if (!VerificaCaso()) return;
      errorProvider.SetError(cmbTipoCaso, " ");
			string stSelect ="Insert Into P_Caso (idTipoCaso"
				+ ", Caso,idcompra "
				+ ") VALUES (" + cmbTipoCaso.Value.ToString()
				+ ", '" + VCaso 
				+ "','" + compra 
				+ "')";
	            Funcion.EjecutaSQL(cdsSeteoF,stSelect);
			try
					{
						MessageBox.Show("Casos Enlazados Correctamente", 
							"Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
					}
					catch (System.Exception ex)
					{
						MessageBox.Show(ex.Message);
					}
			this.Close();
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void cdsCaso_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
						this.cdsCaso.Schema.Connections[0].ConnectionString=
					
							miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			using (TipoCaso miImagen = new TipoCaso())
			{
				miImagen.ShowDialog();
			}
			this.cdsTipoCaso.Clear();  
			this.cdsTipoCaso.Fill();  
			this.cmbTipoCaso.Refresh(); 
		}

		private bool VerificaCaso()
		{
			bool bOk = true;
			string stSelect = string.Format("Exec CompraVerificaCaso {0}, {1}, {2}", 
				VCaso, cmbTipoCaso.Value.ToString(),compra);
			stSelect = Funcion.sEscalarSQL(cdsSeteoF, stSelect);
			if (stSelect.Length > 0)
			{
				bOk = false;
				errorProvider.SetError(cmbTipoCaso, stSelect);
			}		
			return bOk;
		  
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeteoF.Schema.Connections[0].ConnectionString=
					
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void ultraGrid1_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (DialogResult.No ==	MessageBox.Show("¿Desea Borrar Líneas Seleccionadas?",
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
			{
				e.Cancel = true;
			}
			else
			{
				for (int i=0; i<e.Rows.Length; i++)
				{
					string stIdCaso = e.Rows[i].Cells["idCaso"].Value.ToString();
				 	string stSelect="Delete P_Caso where idCaso="+stIdCaso;
					string stSelect1="Select tp.TipoCaso,c.idCaso,c.Caso,c.idtipoCaso from P_Caso c,P_TipoCaso tp where c.idtipocaso=tp.idtipocaso and Caso="+VCaso;
					ultraGrid1.DataSource=Funcion.dvProcedimiento(cdsSeteoF,stSelect1);
					ultraGrid1.Rows.Band.Columns["idCaso"].Hidden=true;
					ultraGrid1.Rows.Band.Columns["Caso"].Hidden=true;
					ultraGrid1.Rows.Band.Columns["idTipoCaso"].Hidden=true;
					string stMensaje = "Tipo de Caso Asignado ha sido Eliminado";
						try
						{
							Funcion.EjecutaSQL(cdsSeteoF,stSelect, true);
							if (stMensaje.Length > 0)
							{
								MessageBox.Show(stMensaje, "Información");
								e.Cancel = true;
								return;
							}
						}
						catch (System.Exception ex)
						{
							MessageBox.Show(ex.Message);
						}
					}
				}
		 }
		
	}
}
