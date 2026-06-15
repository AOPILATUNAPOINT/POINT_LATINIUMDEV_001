using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de DiferenciasEasyMaxLatinium.
	/// </summary>
	public class DiferenciasEasyMaxLatinium : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label55;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		
		string LBodega="";
		DateTime LFecha=DateTime.Now;

		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdFCEasyMax;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdFCLatinium;
		private System.Windows.Forms.Label lblNomBodega;
		private System.Windows.Forms.Label lblFecha;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.TextBox txtRegEasyMax;
		private System.Windows.Forms.TextBox txtRGLatinium;

		private System.ComponentModel.Container components = null;

		public DiferenciasEasyMaxLatinium(string laBodega, DateTime laFecha)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			LBodega = laBodega;
			LFecha = laFecha;
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
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			this.grdFCEasyMax = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.grdFCLatinium = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.label55 = new System.Windows.Forms.Label();
			this.lblNomBodega = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.lblFecha = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label1 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.txtRegEasyMax = new System.Windows.Forms.TextBox();
			this.txtRGLatinium = new System.Windows.Forms.TextBox();
			((System.ComponentModel.ISupportInitialize)(this.grdFCEasyMax)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdFCLatinium)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// grdFCEasyMax
			// 
			this.grdFCEasyMax.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdFCEasyMax.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.grdFCEasyMax.DisplayLayout.Appearance = appearance1;
			this.grdFCEasyMax.DisplayLayout.AutoFitColumns = true;
			this.grdFCEasyMax.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.grdFCEasyMax.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.grdFCEasyMax.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.grdFCEasyMax.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdFCEasyMax.DisplayLayout.Override.HeaderAppearance = appearance3;
			this.grdFCEasyMax.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdFCEasyMax.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdFCEasyMax.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.grdFCEasyMax.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.grdFCEasyMax.Location = new System.Drawing.Point(8, 64);
			this.grdFCEasyMax.Name = "grdFCEasyMax";
			this.grdFCEasyMax.Size = new System.Drawing.Size(960, 184);
			this.grdFCEasyMax.TabIndex = 9;
			this.grdFCEasyMax.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.grdFCEasyMax_InitializeRow);
			// 
			// grdFCLatinium
			// 
			this.grdFCLatinium.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdFCLatinium.Cursor = System.Windows.Forms.Cursors.Default;
			appearance6.BackColor = System.Drawing.Color.White;
			this.grdFCLatinium.DisplayLayout.Appearance = appearance6;
			this.grdFCLatinium.DisplayLayout.AutoFitColumns = true;
			this.grdFCLatinium.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.grdFCLatinium.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.grdFCLatinium.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance7.BackColor = System.Drawing.Color.Transparent;
			this.grdFCLatinium.DisplayLayout.Override.CardAreaAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 10F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdFCLatinium.DisplayLayout.Override.HeaderAppearance = appearance8;
			this.grdFCLatinium.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdFCLatinium.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdFCLatinium.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.grdFCLatinium.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.grdFCLatinium.Location = new System.Drawing.Point(8, 288);
			this.grdFCLatinium.Name = "grdFCLatinium";
			this.grdFCLatinium.Size = new System.Drawing.Size(960, 184);
			this.grdFCLatinium.TabIndex = 10;
			this.grdFCLatinium.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.grdFCLatinium_InitializeRow);
			// 
			// label55
			// 
			this.label55.AutoSize = true;
			this.label55.BackColor = System.Drawing.Color.Transparent;
			this.label55.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label55.Location = new System.Drawing.Point(16, 16);
			this.label55.Name = "label55";
			this.label55.Size = new System.Drawing.Size(54, 16);
			this.label55.TabIndex = 168;
			this.label55.Text = "BODEGA";
			this.label55.TextAlign = System.Drawing.ContentAlignment.TopRight;
			// 
			// lblNomBodega
			// 
			this.lblNomBodega.BackColor = System.Drawing.Color.Transparent;
			this.lblNomBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblNomBodega.Location = new System.Drawing.Point(88, 16);
			this.lblNomBodega.Name = "lblNomBodega";
			this.lblNomBodega.Size = new System.Drawing.Size(400, 16);
			this.lblNomBodega.TabIndex = 169;
			this.lblNomBodega.Text = "Nombre Bodega";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(696, 16);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(37, 16);
			this.label2.TabIndex = 170;
			this.label2.Text = "Fecha";
			this.label2.TextAlign = System.Drawing.ContentAlignment.TopRight;
			// 
			// lblFecha
			// 
			this.lblFecha.BackColor = System.Drawing.Color.Transparent;
			this.lblFecha.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblFecha.Location = new System.Drawing.Point(752, 16);
			this.lblFecha.Name = "lblFecha";
			this.lblFecha.Size = new System.Drawing.Size(112, 16);
			this.lblFecha.TabIndex = 171;
			this.lblFecha.Text = "Fecha";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label4.ForeColor = System.Drawing.SystemColors.HotTrack;
			this.label4.Location = new System.Drawing.Point(8, 34);
			this.label4.Name = "label4";
			this.label4.TabIndex = 172;
			this.label4.Text = "EasyMAX";
			this.label4.TextAlign = System.Drawing.ContentAlignment.TopRight;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label5.ForeColor = System.Drawing.Color.Red;
			this.label5.Location = new System.Drawing.Point(8, 261);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(88, 26);
			this.label5.TabIndex = 173;
			this.label5.Text = "Latinium";
			this.label5.TextAlign = System.Drawing.ContentAlignment.TopRight;
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
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(760, 258);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(102, 16);
			this.label1.TabIndex = 174;
			this.label1.Text = "Total REGISTROS";
			this.label1.TextAlign = System.Drawing.ContentAlignment.TopRight;
			// 
			// label3
			// 
			this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.Location = new System.Drawing.Point(760, 488);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(102, 16);
			this.label3.TabIndex = 175;
			this.label3.Text = "Total REGISTROS";
			this.label3.TextAlign = System.Drawing.ContentAlignment.TopRight;
			// 
			// txtRegEasyMax
			// 
			this.txtRegEasyMax.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtRegEasyMax.Enabled = false;
			this.txtRegEasyMax.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtRegEasyMax.Location = new System.Drawing.Point(864, 256);
			this.txtRegEasyMax.Name = "txtRegEasyMax";
			this.txtRegEasyMax.ReadOnly = true;
			this.txtRegEasyMax.Size = new System.Drawing.Size(64, 20);
			this.txtRegEasyMax.TabIndex = 176;
			this.txtRegEasyMax.Text = "";
			this.txtRegEasyMax.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			// 
			// txtRGLatinium
			// 
			this.txtRGLatinium.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtRGLatinium.Enabled = false;
			this.txtRGLatinium.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtRGLatinium.Location = new System.Drawing.Point(864, 486);
			this.txtRGLatinium.Name = "txtRGLatinium";
			this.txtRGLatinium.ReadOnly = true;
			this.txtRGLatinium.Size = new System.Drawing.Size(64, 20);
			this.txtRGLatinium.TabIndex = 177;
			this.txtRGLatinium.Text = "";
			this.txtRGLatinium.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			// 
			// DiferenciasEasyMaxLatinium
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(976, 517);
			this.Controls.Add(this.txtRGLatinium);
			this.Controls.Add(this.txtRegEasyMax);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.lblFecha);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.lblNomBodega);
			this.Controls.Add(this.label55);
			this.Controls.Add(this.grdFCLatinium);
			this.Controls.Add(this.grdFCEasyMax);
			this.Name = "DiferenciasEasyMaxLatinium";
			this.Text = "DiferenciasEasyMaxLatinium";
			this.Load += new System.EventHandler(this.DiferenciasEasyMaxLatinium_Load);
			((System.ComponentModel.ISupportInitialize)(this.grdFCEasyMax)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdFCLatinium)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		C1.Data.C1DataRow drSeteoF;

		private void DiferenciasEasyMaxLatinium_Load(object sender, System.EventArgs e)
		{
			//Facturas_EasyMax
			//Facturas_Latinium
			Cursor.Current = Cursors.WaitCursor;

			string sSQL = string.Format("Select Nombre From Bodega Where Codigo = '{0}'", LBodega);
			this.lblNomBodega.Text = Funcion.sEscalarSQL(cdsSeteoF, sSQL).ToString().Trim();
			this.lblFecha.Text = LFecha.ToString().Substring(0,11);

			string stExec = string.Format("Exec Facturas_EasyMax '{0}', '{1}'", LBodega, LFecha.ToString("yyyyMMdd"));
			grdFCEasyMax.DisplayLayout.Bands[0].Reset();
			grdFCEasyMax.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stExec);
			this.txtRegEasyMax.Text = this.grdFCEasyMax.Rows.Count.ToString();

			string st2Exec = string.Format("Exec Facturas_Latinium '{0}', '{1}'", LBodega, LFecha.ToString("yyyyMMdd"));
			grdFCLatinium.DisplayLayout.Bands[0].Reset();
			grdFCLatinium.DataSource = Funcion.dvProcedimiento(cdsSeteoF, st2Exec);	
			this.txtRGLatinium.Text = this.grdFCLatinium.Rows.Count.ToString();

			Cursor.Current = Cursors.Default;
		
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		

		}

		private void grdFCLatinium_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			if (e.Row.Cells["CodClte"].Value.ToString() == "ANULADO")
			{
				e.Row.Appearance.BackColor = Color.Yellow;
				e.Row.Appearance.ForeColor = Color.Red;
			}
		}

		private void grdFCEasyMax_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			if (e.Row.Cells["ESTADO"].Value.ToString().Trim() == "ANULADO" || e.Row.Cells["NOMCLI"].Value.ToString().Trim() == "ANULADO")
			{
				e.Row.Appearance.BackColor = Color.Yellow;
				e.Row.Appearance.ForeColor = Color.Red;
			}

		}
	}
}
