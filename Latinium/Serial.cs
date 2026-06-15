using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for Serial.
	/// </summary>
	public class Serial : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraLabel ultraLabel4;
		private Infragistics.Win.Misc.UltraLabel lblContador;
		private Infragistics.Win.Misc.UltraLabel ultraLabel3;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSerie;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCantidad;
		private Infragistics.Win.Misc.UltraButton btGeneraSerial;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtArticulo;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private Infragistics.Win.Misc.UltraButton btConsultar;
		private Infragistics.Win.Misc.UltraButton btGrabar;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Serial()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
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
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
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
			this.ultraLabel4 = new Infragistics.Win.Misc.UltraLabel();
			this.txtCantidad = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblContador = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel3 = new Infragistics.Win.Misc.UltraLabel();
			this.btGeneraSerial = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.btGrabar = new Infragistics.Win.Misc.UltraButton();
			this.txtSerie = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.txtArticulo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btConsultar = new Infragistics.Win.Misc.UltraButton();
			((System.ComponentModel.ISupportInitialize)(this.txtCantidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerie)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraLabel4
			// 
			this.ultraLabel4.AutoSize = true;
			this.ultraLabel4.Location = new System.Drawing.Point(8, 75);
			this.ultraLabel4.Name = "ultraLabel4";
			this.ultraLabel4.Size = new System.Drawing.Size(58, 15);
			this.ultraLabel4.TabIndex = 175;
			this.ultraLabel4.Text = "CANTIDAD";
			// 
			// txtCantidad
			// 
			this.txtCantidad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCantidad.Location = new System.Drawing.Point(80, 71);
			this.txtCantidad.MinValue = 0;
			this.txtCantidad.Name = "txtCantidad";
			this.txtCantidad.PromptChar = ' ';
			this.txtCantidad.Size = new System.Drawing.Size(56, 22);
			this.txtCantidad.TabIndex = 174;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.Location = new System.Drawing.Point(532, 15);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 0);
			this.lblContador.TabIndex = 173;
			// 
			// ultraLabel3
			// 
			this.ultraLabel3.AutoSize = true;
			this.ultraLabel3.Location = new System.Drawing.Point(144, 75);
			this.ultraLabel3.Name = "ultraLabel3";
			this.ultraLabel3.Size = new System.Drawing.Size(41, 15);
			this.ultraLabel3.TabIndex = 172;
			this.ultraLabel3.Text = "SERIAL";
			// 
			// btGeneraSerial
			// 
			this.btGeneraSerial.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btGeneraSerial.Location = new System.Drawing.Point(192, 71);
			this.btGeneraSerial.Name = "btGeneraSerial";
			this.btGeneraSerial.Size = new System.Drawing.Size(112, 26);
			this.btGeneraSerial.TabIndex = 171;
			this.btGeneraSerial.Text = "Genera Seriales";
			this.btGeneraSerial.Visible = false;
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
			this.btCancelar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(220, 432);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(112, 24);
			this.btCancelar.TabIndex = 165;
			this.btCancelar.Text = "&Salir";
			this.btCancelar.Visible = false;
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btGrabar
			// 
			this.btGrabar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btGrabar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btGrabar.Location = new System.Drawing.Point(100, 432);
			this.btGrabar.Name = "btGrabar";
			this.btGrabar.Size = new System.Drawing.Size(112, 24);
			this.btGrabar.TabIndex = 164;
			this.btGrabar.Text = "&Grabar";
			this.btGrabar.Visible = false;
			// 
			// txtSerie
			// 
			this.txtSerie.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtSerie.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSerie.Location = new System.Drawing.Point(200, 71);
			this.txtSerie.Name = "txtSerie";
			this.txtSerie.Size = new System.Drawing.Size(344, 22);
			this.txtSerie.TabIndex = 161;
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
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance2;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
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
			this.ultraGrid1.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(0, 103);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(552, 321);
			this.ultraGrid1.TabIndex = 157;
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(8, 40);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(57, 15);
			this.ultraLabel1.TabIndex = 154;
			this.ultraLabel1.Text = "ARTÍCULO";
			// 
			// txtArticulo
			// 
			appearance6.ForeColor = System.Drawing.Color.Black;
			this.txtArticulo.Appearance = appearance6;
			this.txtArticulo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtArticulo.Enabled = false;
			this.txtArticulo.Location = new System.Drawing.Point(80, 40);
			this.txtArticulo.Name = "txtArticulo";
			this.txtArticulo.Size = new System.Drawing.Size(464, 22);
			this.txtArticulo.TabIndex = 153;
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(8, 8);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(46, 15);
			this.ultraLabel2.TabIndex = 177;
			this.ultraLabel2.Text = "CÓDIGO";
			// 
			// txtCodigo
			// 
			appearance7.BackColor = System.Drawing.Color.White;
			appearance7.ForeColor = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance7;
			this.txtCodigo.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.txtCodigo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Enabled = false;
			this.txtCodigo.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtCodigo.ForeColor = System.Drawing.Color.Black;
			this.txtCodigo.Location = new System.Drawing.Point(80, 8);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(464, 22);
			this.txtCodigo.TabIndex = 176;
			// 
			// btConsultar
			// 
			this.btConsultar.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
			this.btConsultar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btConsultar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btConsultar.Location = new System.Drawing.Point(340, 432);
			this.btConsultar.Name = "btConsultar";
			this.btConsultar.Size = new System.Drawing.Size(112, 24);
			this.btConsultar.TabIndex = 178;
			this.btConsultar.Text = "&Consultar";
			this.btConsultar.Visible = false;
			// 
			// Serial
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(552, 462);
			this.Controls.Add(this.btConsultar);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.ultraLabel4);
			this.Controls.Add(this.txtCantidad);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.ultraLabel3);
			this.Controls.Add(this.btGeneraSerial);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btGrabar);
			this.Controls.Add(this.txtSerie);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.txtArticulo);
			this.Name = "Serial";
			this.Text = "Serial";
			this.Load += new System.EventHandler(this.Serial_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtCantidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerie)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void Serial_Load(object sender, System.EventArgs e)
		{
		
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}
	}
}

