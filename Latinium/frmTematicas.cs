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
using System.Text;
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;
using System.Security.Cryptography.Xml;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmTematicas.
	/// </summary>
	public class frmTematicas : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnGrabar;
		private C1.Data.C1DataSet cdsBodega;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugTematicas;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmTematicas()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTematica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTematica");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmTematicas));
			this.ugTematicas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnGrabar = new System.Windows.Forms.Button();
			this.cdsBodega = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.ugTematicas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsBodega)).BeginInit();
			this.SuspendLayout();
			// 
			// ugTematicas
			// 
			this.ugTematicas.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugTematicas.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugTematicas.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ugTematicas.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 300;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.ugTematicas.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugTematicas.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.ugTematicas.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ugTematicas.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ugTematicas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.ugTematicas.DisplayLayout.Override.CardAreaAppearance = appearance3;
			this.ugTematicas.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 10F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugTematicas.DisplayLayout.Override.HeaderAppearance = appearance4;
			appearance5.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance5.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugTematicas.DisplayLayout.Override.RowAlternateAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugTematicas.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugTematicas.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.ugTematicas.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugTematicas.Location = new System.Drawing.Point(8, 8);
			this.ugTematicas.Name = "ugTematicas";
			this.ugTematicas.Size = new System.Drawing.Size(328, 392);
			this.ugTematicas.TabIndex = 19;
			this.ugTematicas.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ugTematicas_AfterRowInsert);
			this.ugTematicas.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ugTematicas_BeforeRowsDeleted);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// btnGrabar
			// 
			this.btnGrabar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGrabar.CausesValidation = false;
			this.btnGrabar.Image = ((System.Drawing.Image)(resources.GetObject("btnGrabar.Image")));
			this.btnGrabar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGrabar.Location = new System.Drawing.Point(8, 408);
			this.btnGrabar.Name = "btnGrabar";
			this.btnGrabar.Size = new System.Drawing.Size(72, 32);
			this.btnGrabar.TabIndex = 20;
			this.btnGrabar.Text = "Grabar";
			this.btnGrabar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGrabar.Click += new System.EventHandler(this.btnGrabar_Click);
			// 
			// cdsBodega
			// 
			this.cdsBodega.BindingContextCtrl = this;
			this.cdsBodega.DataLibrary = "LibFacturacion";
			this.cdsBodega.DataLibraryUrl = "";
			this.cdsBodega.DataSetDef = "dsBodega";
			this.cdsBodega.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsBodega.Name = "cdsBodega";
			this.cdsBodega.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsBodega.SchemaDef = null;
			this.cdsBodega.BeforeFill += new C1.Data.FillEventHandler(this.cdsBodega_BeforeFill);
			// 
			// frmTematicas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(344, 445);
			this.Controls.Add(this.btnGrabar);
			this.Controls.Add(this.ugTematicas);
			this.Name = "frmTematicas";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Registro de TEMATICAS";
			this.Load += new System.EventHandler(this.frmTematicas_Load);
			((System.ComponentModel.ISupportInitialize)(this.ugTematicas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsBodega)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsBodega_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsBodega.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void VisulizaTematicas()
		{
			string lcQry = "Select idTematica,Nombre From Tematica Order By Nombre";
			this.ugTematicas.DataSource = Funcion.dvProcedimiento(cdsBodega, lcQry);
		}

		private void frmTematicas_Load(object sender, System.EventArgs e)
		{
			VisulizaTematicas();
		}

		private void btnGrabar_Click(object sender, System.EventArgs e)
		{
			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
			{
				oConexion.Open();

				SqlTransaction oTransaction = null;
				SqlCommand oCmdActualiza = oConexion.CreateCommand();
				oCmdActualiza.Connection = oConexion;
				oCmdActualiza.CommandType = CommandType.Text;
				oCmdActualiza.CommandTimeout = 0;
						
				try
				{							
					oTransaction = oConexion.BeginTransaction();
					oCmdActualiza.Transaction = oTransaction;

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ugTematicas.Rows.All)
					{
						string sSQLDetalle = string.Format("Exec GrabaTematicas {0}, '{1}'",
							(int)dr.Cells["idTematica"].Value, 
							dr.Cells["Nombre"].Value.ToString());
						oCmdActualiza.CommandText = sSQLDetalle;
						oCmdActualiza.ExecuteNonQuery();
					}
					oTransaction.Commit();
				}
				catch (Exception ex)
				{
					try
					{
						oTransaction.Rollback();

						MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
					catch (Exception ex2)
					{
						MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);							
					}
				}
				finally
				{
					oConexion.Close();
				}
			}
			// Visuliza información
			MessageBox.Show("Información grabada correctamente...!!!", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Information);
			VisulizaTematicas();
		}

		private void ugTematicas_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idTematica"].Value = 0;
			e.Row.Cells["Nombre"].Value = "";
		}

		private void ugTematicas_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (DialogResult.No != MessageBox.Show("¿Está Seguro de Eliminar Temática?", "Información...!!!", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
			{
				for (int i=0; i<e.Rows.Length; i++)
				{					
					string sSQL = string.Format("Delete From Tematica Where idTematica = {0}", (int)e.Rows[i].Cells["idTematica"].Value);
					Funcion.EjecutaSQL(cdsBodega, sSQL, true);
					//MessageBox.Show("Serial Eliminada del Conteo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
				e.DisplayPromptMsg = false;
			}
			else  e.Cancel = true;	
			
		}
	}
}
