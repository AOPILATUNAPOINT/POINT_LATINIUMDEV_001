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
	/// Descripción breve de CompraSaldoArt.
	/// </summary>
	public class CompraSaldoArt : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private C1.Data.C1DataSet cdsGrupo;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		int IdArticulo = 0;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label lblArt;
		public CompraSaldoArt(int idArticulo)
		{
			InitializeComponent();
			IdArticulo = idArticulo;
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
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.cdsGrupo = new C1.Data.C1DataSet();
			this.lblArt = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupo)).BeginInit();
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
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowColSwapping = Infragistics.Win.UltraWinGrid.AllowColSwapping.NotAllowed;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
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
			this.ultraGrid1.Location = new System.Drawing.Point(24, 40);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(466, 168);
			this.ultraGrid1.TabIndex = 3;
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.Location = new System.Drawing.Point(208, 224);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(80, 24);
			this.btAceptar.TabIndex = 4;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// cdsGrupo
			// 
			this.cdsGrupo.BindingContextCtrl = this;
			this.cdsGrupo.DataLibrary = "LibFacturacion";
			this.cdsGrupo.DataLibraryUrl = "";
			this.cdsGrupo.DataSetDef = "dsSeteoF";
			this.cdsGrupo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsGrupo.Name = "cdsGrupo";
			this.cdsGrupo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsGrupo.SchemaDef = null;
			// 
			// lblArt
			// 
			this.lblArt.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.lblArt.Location = new System.Drawing.Point(80, 16);
			this.lblArt.Name = "lblArt";
			this.lblArt.Size = new System.Drawing.Size(400, 16);
			this.lblArt.TabIndex = 5;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(24, 16);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(45, 16);
			this.label2.TabIndex = 6;
			this.label2.Text = "Articulo:";
			// 
			// CompraSaldoArt
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(512, 266);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.lblArt);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.ultraGrid1);
			this.Name = "CompraSaldoArt";
			this.Text = "Saldo por Bodega";
			this.Load += new System.EventHandler(this.CompraSaldoArt_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			Close();
		}
		DataSet miDataSet = new DataSet();
		DataView dv = new DataView();
		DataSet miDataSetG = new DataSet();
		DataView dvG = new DataView();
		DataSet miDataSetT = new DataSet();
		DataView dvT = new DataView();

		private void CompraSaldoArt_Load(object sender, System.EventArgs e)
		{
			lblArt.Text = Funcion.sEscalarSQL(cdsGrupo, "Select Articulo From Articulo Where idArticulo = " + IdArticulo);
			Funcion miFuncion = new Funcion();
			string strConn = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
			string stProcedimiento = "exec ArticuloSaldoBodega " + IdArticulo.ToString();
			SqlDataAdapter myData = new SqlDataAdapter(stProcedimiento, strConn);

			myData.TableMappings.Add("Tabla", "ExcelSube");
			
			miDataSetG.Reset();
			miDataSetG.Clear();
			try
			{
				myData.Fill(miDataSetG);
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al abrir la Hoja A del Archivo: " );
				return;
			}

			dvG = miDataSetG.Tables[0].DefaultView;
			ultraGrid1.DataSource = dvG;
		}
	}
}
