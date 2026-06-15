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
	/// Descripción breve de RolContatoCaduca.
	/// </summary>
	public class RolContatoCaduca : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private C1.Data.C1DataSet cdsContrato;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public RolContatoCaduca()
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
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.cdsContrato = new C1.Data.C1DataSet();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			((System.ComponentModel.ISupportInitialize)(this.cdsContrato)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			this.SuspendLayout();
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.Location = new System.Drawing.Point(232, 296);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(72, 24);
			this.btAceptar.TabIndex = 5;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btCancelar.Location = new System.Drawing.Point(392, 296);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(72, 24);
			this.btCancelar.TabIndex = 6;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Visible = false;
			// 
			// cdsContrato
			// 
			this.cdsContrato.BindingContextCtrl = this;
			this.cdsContrato.DataLibrary = "LibPersonal";
			this.cdsContrato.DataLibraryUrl = "";
			this.cdsContrato.DataSetDef = "dsPersonal";
			this.cdsContrato.FillOnRequest = false;
			this.cdsContrato.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsContrato.Name = "cdsContrato";
			this.cdsContrato.SchemaClassName = "LibPersonal.DataClass";
			this.cdsContrato.SchemaDef = null;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			this.ultraGrid1.Location = new System.Drawing.Point(24, 24);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(488, 248);
			this.ultraGrid1.TabIndex = 7;
			// 
			// RolContatoCaduca
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(536, 333);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Name = "RolContatoCaduca";
			this.Text = "Contatos Caducados";
			((System.ComponentModel.ISupportInitialize)(this.cdsContrato)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		DataSet miDataSetG = new DataSet();
		DataView dvG = new DataView();

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			Funcion miFuncion = new Funcion();
			string strConn = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
			string stProcedimiento = "exec RolCaduca " ;
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
