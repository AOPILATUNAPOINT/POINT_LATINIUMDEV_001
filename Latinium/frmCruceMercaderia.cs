using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCruceMercaderia.
	/// </summary>
	public class frmCruceMercaderia : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtpHora;
		private System.Windows.Forms.Label lblHora;
		private System.Windows.Forms.DateTimePicker txtFecha;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.TextBox txtNSerie;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.TextBox txtCodArt;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label lblBodega;
		private System.Windows.Forms.Button button2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.ToolTip toolTip1;
		private System.ComponentModel.IContainer components;

		string lBodeg = "";
		DateTime FeToma = DateTime.Today;
		DateTime HoToma = DateTime.Now;
		int Ar1 = 0;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraGrid grContInv;
		private System.Windows.Forms.TextBox txtObservaciones;
		int Ar2 = 0;
		int ArtCam = 0;
		int idCnt = 0;

		public frmCruceMercaderia(string laBodega, DateTime FechaToma, DateTime HoraToma, int idArti1, int idArti2)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			lBodeg = laBodega;
			FeToma = FechaToma;
			HoToma = HoraToma;
			Ar1 = idArti1;
			Ar2 = idArti2;

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
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCruceMercaderia));
			this.dtpHora = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.lblHora = new System.Windows.Forms.Label();
			this.txtFecha = new System.Windows.Forms.DateTimePicker();
			this.label4 = new System.Windows.Forms.Label();
			this.grContInv = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.txtNSerie = new System.Windows.Forms.TextBox();
			this.label2 = new System.Windows.Forms.Label();
			this.txtCodArt = new System.Windows.Forms.TextBox();
			this.label1 = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.button2 = new System.Windows.Forms.Button();
			this.txtObservaciones = new System.Windows.Forms.TextBox();
			this.label3 = new System.Windows.Forms.Label();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.dtpHora)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grContInv)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// dtpHora
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtpHora.Appearance = appearance1;
			this.dtpHora.DateTime = new System.DateTime(1753, 1, 1, 0, 0, 0, 0);
			this.dtpHora.DropDownButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Never;
			this.dtpHora.Enabled = false;
			this.dtpHora.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.dtpHora.FormatString = "HH:mm";
			this.dtpHora.Location = new System.Drawing.Point(272, 16);
			this.dtpHora.MaskInput = "{LOC}hh:mm";
			this.dtpHora.Name = "dtpHora";
			this.dtpHora.Size = new System.Drawing.Size(104, 28);
			this.dtpHora.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.dtpHora.TabIndex = 249;
			this.dtpHora.Value = null;
			// 
			// lblHora
			// 
			this.lblHora.AutoSize = true;
			this.lblHora.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblHora.Location = new System.Drawing.Point(216, 19);
			this.lblHora.Name = "lblHora";
			this.lblHora.Size = new System.Drawing.Size(43, 22);
			this.lblHora.TabIndex = 248;
			this.lblHora.Text = "Hora";
			this.lblHora.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtFecha
			// 
			this.txtFecha.CustomFormat = "";
			this.txtFecha.Enabled = false;
			this.txtFecha.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtFecha.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.txtFecha.Location = new System.Drawing.Point(72, 17);
			this.txtFecha.Name = "txtFecha";
			this.txtFecha.Size = new System.Drawing.Size(136, 26);
			this.txtFecha.TabIndex = 245;
			this.txtFecha.Value = new System.DateTime(2014, 6, 23, 0, 0, 0, 0);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label4.Location = new System.Drawing.Point(8, 19);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(58, 22);
			this.label4.TabIndex = 244;
			this.label4.Text = "Fecha:";
			// 
			// grContInv
			// 
			this.grContInv.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grContInv.Cursor = System.Windows.Forms.Cursors.Default;
			appearance2.BackColor = System.Drawing.Color.White;
			this.grContInv.DisplayLayout.Appearance = appearance2;
			this.grContInv.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.grContInv.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 10F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grContInv.DisplayLayout.Override.HeaderAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grContInv.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grContInv.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.grContInv.Location = new System.Drawing.Point(8, 96);
			this.grContInv.Name = "grContInv";
			this.grContInv.Size = new System.Drawing.Size(576, 288);
			this.grContInv.TabIndex = 242;
			this.grContInv.DoubleClick += new System.EventHandler(this.grContInv_DoubleClick);
			// 
			// txtNSerie
			// 
			this.txtNSerie.BackColor = System.Drawing.SystemColors.Window;
			this.txtNSerie.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNSerie.Location = new System.Drawing.Point(104, 416);
			this.txtNSerie.Name = "txtNSerie";
			this.txtNSerie.Size = new System.Drawing.Size(384, 20);
			this.txtNSerie.TabIndex = 237;
			this.txtNSerie.Text = "";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(40, 418);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(59, 16);
			this.label2.TabIndex = 240;
			this.label2.Text = "# de Serie:";
			// 
			// txtCodArt
			// 
			this.txtCodArt.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCodArt.Location = new System.Drawing.Point(104, 390);
			this.txtCodArt.Name = "txtCodArt";
			this.txtCodArt.Size = new System.Drawing.Size(224, 20);
			this.txtCodArt.TabIndex = 238;
			this.txtCodArt.Text = "";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(54, 392);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(45, 16);
			this.label1.TabIndex = 239;
			this.label1.Text = "Artículo:";
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblBodega.ForeColor = System.Drawing.Color.ForestGreen;
			this.lblBodega.Location = new System.Drawing.Point(8, 56);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(97, 25);
			this.lblBodega.TabIndex = 250;
			this.lblBodega.Text = "BODEGA:";
			// 
			// button2
			// 
			this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
			this.button2.Location = new System.Drawing.Point(16, 496);
			this.button2.Name = "button2";
			this.button2.Size = new System.Drawing.Size(48, 40);
			this.button2.TabIndex = 251;
			this.toolTip1.SetToolTip(this.button2, "Procesar INFORMACION");
			this.button2.Click += new System.EventHandler(this.button2_Click);
			// 
			// txtObservaciones
			// 
			this.txtObservaciones.BackColor = System.Drawing.SystemColors.Window;
			this.txtObservaciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObservaciones.Location = new System.Drawing.Point(104, 437);
			this.txtObservaciones.Multiline = true;
			this.txtObservaciones.Name = "txtObservaciones";
			this.txtObservaciones.Size = new System.Drawing.Size(448, 51);
			this.txtObservaciones.TabIndex = 252;
			this.txtObservaciones.Text = "";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(16, 440);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(83, 16);
			this.label3.TabIndex = 253;
			this.label3.Text = "Observaciones:";
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
			// frmCruceMercaderia
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(592, 541);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.txtObservaciones);
			this.Controls.Add(this.button2);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.dtpHora);
			this.Controls.Add(this.lblHora);
			this.Controls.Add(this.txtFecha);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.grContInv);
			this.Controls.Add(this.txtNSerie);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtCodArt);
			this.Controls.Add(this.label1);
			this.Name = "frmCruceMercaderia";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "frmCruceMercaderia";
			this.Load += new System.EventHandler(this.frmCruceMercaderia_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtpHora)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grContInv)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		C1.Data.C1DataRow drSeteoF;

		private void frmCruceMercaderia_Load(object sender, System.EventArgs e)
		{
			this.txtFecha.Value = FeToma; /*DateTime.Today;*/
			this.dtpHora.Value = HoToma;
			string sSQL = string.Format("Select Nombre From Bodega Where Codigo = '{0}'", lBodeg);
			this.lblBodega.Text = "BODEGA: "+Funcion.sEscalarSQL(cdsSeteoF, sSQL).ToString().Trim();
			string lcQry = string.Format("Exec Arti_Cruce '{0}','{1}',{2},{3}",lBodeg,FeToma.ToString("yyyyMMdd"),Ar1,Ar2);
			grContInv.DataSource = Funcion.dvProcedimiento(cdsSeteoF, lcQry);	
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void grContInv_DoubleClick(object sender, System.EventArgs e)
		{
			/* Recupera información para procesar */
			idCnt = (int)this.grContInv.ActiveRow.Cells["idConteo"].Value;
			int idArtclo = (int)this.grContInv.ActiveRow.Cells["idArticulo"].Value;
			this.txtNSerie.Text = this.grContInv.ActiveRow.Cells["NumSerieArt"].Value.ToString();
			if (idArtclo==Ar1)
			{
				string Qry = string.Format("Select Codigo From Articulo Where idArticulo = {0}", Ar2);
				this.txtCodArt.Text = Funcion.sEscalarSQL(cdsSeteoF, Qry).ToString().Trim();
				ArtCam = Ar2;
				this.txtObservaciones.Focus();
			}
			else
			{
				string Qry = string.Format("Select Codigo From Articulo Where idArticulo = {0}", Ar1);
				this.txtCodArt.Text = Funcion.sEscalarSQL(cdsSeteoF, Qry).ToString().Trim();
				ArtCam = Ar1;
				this.txtObservaciones.Focus();
			}
		}

		private void button2_Click(object sender, System.EventArgs e)
		{
			Cursor.Current = Cursors.WaitCursor;
			
  		string stExec = string.Format("Exec Cruce_Mercaderia {0},{1},'{2}'", ArtCam, idCnt, this.txtObservaciones.Text.ToString());
			Funcion.EjecutaSQL(cdsSeteoF, stExec);
			MessageBox.Show("Proceso terminado...", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
			this.DialogResult = DialogResult.OK;
			
			Cursor.Current = Cursors.Default;
		}

	}
}
