using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using CrystalDecisions.Shared;
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;
using System.IO;
using System.Diagnostics;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmLineaCreditoAvaluoAnadir.
	/// </summary>
	public class frmLineaCreditoAvaluoAnadir : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnCargar;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.GroupBox groupBox3;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udgAvaluo;
		private System.Windows.Forms.Button button1;
		/// <summary>
		/// Variable del diseñador requerida.
		int idGarantiaLineaCreditoAvaluoDoc = 0;
		bool bNuevo = false;
		bool bVer = false;
		int idGarantiaLineaCreditoAvaluo = 0;
		bool bEliminar = false;

		string txtDirectorioarchivo = "";
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDocumento;
		private System.Windows.Forms.Button button2;
		private System.Windows.Forms.Button button3;
		private System.Windows.Forms.GroupBox groupBox2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbAvaluo;
		public System.Windows.Forms.Label label1;
		private System.Windows.Forms.Button button4;
		private System.Windows.Forms.Button button5;
		/// </summary>
		private System.ComponentModel.Container components = null;
		int iTipo = 0;

		public frmLineaCreditoAvaluoAnadir(int IdGarantiaLineaCreditoAvaluo, int ITipo)
		{
			iTipo = ITipo;
			idGarantiaLineaCreditoAvaluo = IdGarantiaLineaCreditoAvaluo;
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idListadoAvaluo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmLineaCreditoAvaluoAnadir));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idListadoAvaluo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			this.udgAvaluo = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnCargar = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.button1 = new System.Windows.Forms.Button();
			this.txtDocumento = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.button2 = new System.Windows.Forms.Button();
			this.button3 = new System.Windows.Forms.Button();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.cmbAvaluo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.button4 = new System.Windows.Forms.Button();
			this.button5 = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.udgAvaluo)).BeginInit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtDocumento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbAvaluo)).BeginInit();
			this.SuspendLayout();
			// 
			// udgAvaluo
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn1.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			this.udgAvaluo.Band.Columns.AddRange(new object[] {
																													ultraDataColumn1,
																													ultraDataColumn2});
			// 
			// btnCargar
			// 
			this.btnCargar.Location = new System.Drawing.Point(400, 24);
			this.btnCargar.Name = "btnCargar";
			this.btnCargar.Size = new System.Drawing.Size(80, 23);
			this.btnCargar.TabIndex = 1037;
			this.btnCargar.Text = "Cargar";
			this.btnCargar.Click += new System.EventHandler(this.btnCargar_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Controls.Add(this.button1);
			this.groupBox1.Controls.Add(this.txtDocumento);
			this.groupBox1.Controls.Add(this.btnCargar);
			this.groupBox1.Location = new System.Drawing.Point(16, 112);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(584, 64);
			this.groupBox1.TabIndex = 1040;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "* Documento";
			// 
			// button1
			// 
			this.button1.Location = new System.Drawing.Point(496, 24);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(72, 23);
			this.button1.TabIndex = 1040;
			this.button1.Text = "Ver";
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// txtDocumento
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDocumento.Appearance = appearance1;
			this.txtDocumento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDocumento.Enabled = false;
			this.txtDocumento.Location = new System.Drawing.Point(16, 24);
			this.txtDocumento.MaxLength = 50;
			this.txtDocumento.Name = "txtDocumento";
			this.txtDocumento.Size = new System.Drawing.Size(368, 21);
			this.txtDocumento.TabIndex = 1039;
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(936, 56);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(80, 23);
			this.btnSalir.TabIndex = 1042;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(936, 16);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(82, 23);
			this.btnGuardar.TabIndex = 1041;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// groupBox3
			// 
			this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox3.Location = new System.Drawing.Point(-904, 88);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(2840, 8);
			this.groupBox3.TabIndex = 1043;
			this.groupBox3.TabStop = false;
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
			// button2
			// 
			this.button2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button2.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button2.CausesValidation = false;
			this.button2.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.button2.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
			this.button2.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button2.Location = new System.Drawing.Point(776, 56);
			this.button2.Name = "button2";
			this.button2.Size = new System.Drawing.Size(80, 23);
			this.button2.TabIndex = 1046;
			this.button2.Text = "Salir";
			this.button2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// button3
			// 
			this.button3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button3.Image = ((System.Drawing.Image)(resources.GetObject("button3.Image")));
			this.button3.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button3.Location = new System.Drawing.Point(776, 16);
			this.button3.Name = "button3";
			this.button3.Size = new System.Drawing.Size(82, 23);
			this.button3.TabIndex = 1045;
			this.button3.Text = "&Guardar";
			this.button3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button3.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Controls.Add(this.cmbAvaluo);
			this.groupBox2.Controls.Add(this.label1);
			this.groupBox2.Location = new System.Drawing.Point(16, 16);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(480, 64);
			this.groupBox2.TabIndex = 1047;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Datos principales";
			// 
			// cmbAvaluo
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbAvaluo.Appearance = appearance2;
			this.cmbAvaluo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbAvaluo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbAvaluo.DataSource = this.udgAvaluo;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.Caption = "Descripción";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 202;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbAvaluo.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbAvaluo.DisplayMember = "Descripcion";
			this.cmbAvaluo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbAvaluo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbAvaluo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbAvaluo.Location = new System.Drawing.Point(136, 24);
			this.cmbAvaluo.Name = "cmbAvaluo";
			this.cmbAvaluo.Size = new System.Drawing.Size(192, 21);
			this.cmbAvaluo.TabIndex = 817;
			this.cmbAvaluo.ValueMember = "idListadoAvaluo";
			this.cmbAvaluo.ValueChanged += new System.EventHandler(this.cmbAvaluo_ValueChanged);
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(8, 24);
			this.label1.Name = "label1";
			this.label1.TabIndex = 812;
			this.label1.Text = "* Avalúo:";
			// 
			// button4
			// 
			this.button4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button4.CausesValidation = false;
			this.button4.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.button4.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.button4.Image = ((System.Drawing.Image)(resources.GetObject("button4.Image")));
			this.button4.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button4.Location = new System.Drawing.Point(520, 56);
			this.button4.Name = "button4";
			this.button4.Size = new System.Drawing.Size(80, 23);
			this.button4.TabIndex = 1049;
			this.button4.Text = "Salir";
			this.button4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button4.Click += new System.EventHandler(this.button4_Click);
			// 
			// button5
			// 
			this.button5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button5.Image = ((System.Drawing.Image)(resources.GetObject("button5.Image")));
			this.button5.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button5.Location = new System.Drawing.Point(520, 16);
			this.button5.Name = "button5";
			this.button5.Size = new System.Drawing.Size(82, 23);
			this.button5.TabIndex = 1048;
			this.button5.Text = "&Guardar";
			this.button5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button5.Click += new System.EventHandler(this.button5_Click);
			// 
			// frmLineaCreditoAvaluoAnadir
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(608, 190);
			this.Controls.Add(this.button4);
			this.Controls.Add(this.button5);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.button2);
			this.Controls.Add(this.button3);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.groupBox1);
			this.Name = "frmLineaCreditoAvaluoAnadir";
			this.Text = "Añadir avalúo";
			this.Load += new System.EventHandler(this.frmLineaCreditoAvaluoAnadir_Load);
			((System.ComponentModel.ISupportInitialize)(this.udgAvaluo)).EndInit();
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtDocumento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbAvaluo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmLineaCreditoAvaluoAnadir_Load(object sender, System.EventArgs e)
		{

			if(iTipo == frmLineaCredito.iNuevo)
			{
				bNuevo = true;
				this.Text = "Añadir avalúo";
			}
			if(iTipo == frmLineaCredito.iVer)
			{
				bVer = true;
				this.Text = "Ver avalúo";
			}
			if(iTipo == frmLineaCredito.iEliminar)
			{
				bEliminar = true;
				this.Text = "Eliminar avalúo";
			}

			//Cargar avalúo
			this.cmbAvaluo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, String.Format("EXEC LineaCreditoAvaluoCon"));
			LineaCreditoGarantiaAvaluoModel();
			ControlVista();
		}

		private void ControlVista()
		{
			if(bVer)
			{
				this.cmbAvaluo.Enabled = false;
				this.btnCargar.Enabled = false;
				this.button5.Enabled = false;
			}

			if(bEliminar)
			{
				this.cmbAvaluo.Enabled = false;
				this.btnCargar.Enabled = false;
				this.button5.Text = "Eliminar";
			}
		}

		private void LineaCreditoGarantiaAvaluoModel()
		{
			if(idGarantiaLineaCreditoAvaluo != 0)
			{
				this.cmbAvaluo.Value = Funcion.iEscalarSQL(cdsSeteoF, String.Format("EXEC LineaCreditoGarantiaAvaluoModel '{0}', {1}", "idListadoAvaluo", idGarantiaLineaCreditoAvaluo));
				//Visualización del documento y consultar id de doc
				idGarantiaLineaCreditoAvaluoDoc = Funcion.iEscalarSQL(cdsSeteoF, String.Format("Select top 1 idGarantiaLineaCreditoAvaluoDoc From GarantiaLineaCreditoAvaluoDoc where idGarantiaLineaCreditoAvaluo = {0}", idGarantiaLineaCreditoAvaluo));
				bool Cargado = Funcion.bEscalarSQL(cdsSeteoF, String.Format("Select top 1 Cargado From GarantiaLineaCreditoAvaluoDoc where idGarantiaLineaCreditoAvaluo = {0}", idGarantiaLineaCreditoAvaluo), false);
				if(Cargado)
				{
					this.txtDocumento.Text = Funcion.sEscalarSQL(cdsSeteoF, String.Format("EXEC LineaCreditoGarantiaAvaluoDocModel '{0}', {1}", "sCarpeta", idGarantiaLineaCreditoAvaluoDoc));
				}
				else
				{
					this.txtDocumento.Text = Funcion.sEscalarSQL(cdsSeteoF, String.Format("EXEC LineaCreditoGarantiaAvaluoDocModel '{0}', {1}", "txtDirectorioarchivo", idGarantiaLineaCreditoAvaluoDoc));
				}

			}
		}


		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			
		}

		private void btnCargar_Click(object sender, System.EventArgs e)
		{
				if (bNuevo)
				{
					#region Cargar documento 
					try
					{

						#region Cargar documento
						OpenFileDialog buscar = new OpenFileDialog ();
						if (buscar.ShowDialog () == DialogResult.OK)
						{
							//Guardar variable dirección del archivo
							txtDirectorioarchivo = buscar.FileName;

							#region Eliminar archivo anterior
							if(this.txtDocumento.Text != txtDirectorioarchivo && this.txtDocumento.Text != "" && idGarantiaLineaCreditoAvaluoDoc != 0)
							{
								//Eliminar anterior seleciconado
								try
								{
									//Si se guardo crear un registro por el nuevo guardado
									bool bCont =Funcion.bEjecucion(frmRequerimiento.sconexionPoint, string.Format("Exec LineaCreditoGarantiaAvaluoDocEliminar {0}", 
										idGarantiaLineaCreditoAvaluoDoc), false);
									if (bCont)
									{
										//Setear el documento para poder se guardado de nuevo
										this.txtDocumento.Text = "";
									}
								}
								catch (Exception ex)
								{
									string sms = ex.Message;
									return;						
								}
							}
							#endregion Eliminar archivo anterior

							//Sacar el tamaño del archivo subido
							System.IO.FileInfo info = new System.IO.FileInfo(txtDirectorioarchivo);
							//El lenght está en tamaños bytes
							long itamano = info.Length;
							decimal dtamano = Convert.ToDecimal(itamano);
							if(!frmLineaCredito.LineaCreditoConfgVal(cdsSeteoF,frmLineaCredito.iConfiguracionDocumento, dtamano, 1))return;

							//Si se guardo crear un registro por el nuevo guardado
							DataView dvLineaCreditoGarantiaAvaluoDocCrear = new DataView();
							dvLineaCreditoGarantiaAvaluoDocCrear = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec LineaCreditoGarantiaAvaluoDocCrear '{0}', {1}, {2}, {3}", 
								txtDirectorioarchivo, 
								frmLineaCreditoCrear.idLineaCreditoTransaccion,
								0,
								0
								));
							  
							DataSet dsLineaCreditoGarantiaAvaluoDocCrear = new DataSet();
							dsLineaCreditoGarantiaAvaluoDocCrear = dvLineaCreditoGarantiaAvaluoDocCrear.Table.DataSet;
							bool bResultado = bool.Parse(dsLineaCreditoGarantiaAvaluoDocCrear.Tables[0].Rows[0]["bResultado"].ToString());
							string sMensaje = dsLineaCreditoGarantiaAvaluoDocCrear.Tables[0].Rows[0]["sMensaje"].ToString();

							if(bResultado == false)
							{
								MessageBox.Show(string.Format(sMensaje), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
								return;
							}
							else
							{
								this.txtDocumento.Text = txtDirectorioarchivo;
								idGarantiaLineaCreditoAvaluoDoc = Convert.ToInt32(sMensaje);
							}
								
						}
						#endregion Cargar documento
					}	
					catch (Exception ex)
					{
						MessageBox.Show ("Error: "+ ex.Message);
					}
					#endregion Cargar documento
				}
				else
				{
					#region error
					try
					{
						OpenFileDialog buscar = new OpenFileDialog ();
						if (buscar.ShowDialog () == DialogResult.OK)
						{
							txtDirectorioarchivo = buscar.FileName;
						}	
					}
					catch (Exception ex)
					{
						MessageBox.Show ("Error: "+ ex.Message);
					}
					#endregion error
				}
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			try
			{
					txtDirectorioarchivo = this.txtDocumento.Text;
					if (txtDirectorioarchivo.Length > 10 )
					{
						Process.Start(txtDirectorioarchivo);
					}
			}
			catch (Exception ex)
			{
				string sms = ex.Message;
				return;						
			}
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void cmbAvaluo_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void button5_Click(object sender, System.EventArgs e)
		{
			#region Nuevo avaluo
			if (bNuevo)
			{
				#region Valida campos  
				if (!Validacion.vbComboVacio(this.cmbAvaluo, "Ingrese un avalúo")) return;
				#endregion Valida campos  
				//Verificar que se suba el archivo para evidencia del avalúo
				if (this.txtDocumento.Text == "")
				{
					decimal valor = 0;
					if(!frmLineaCredito.LineaCreditoConfgVal(cdsSeteoF, frmLineaCredito.iConfiguracionCargarDocumentoAvaluo, valor, 1)) return;
				}
				int idGarantiaLineaCreditoAvaluo = 0;
				string squery = "";
				try
				{
					//Crear registro avalúo
					squery = String.Format("EXEC LineaCreditoGarantiaAvaluoCrear {0}, {1}", 
						frmLineaCreditoCrear.idLineaCreditoTransaccion,
						(int)this.cmbAvaluo.Value
						);
					idGarantiaLineaCreditoAvaluo = Funcion.iEscalarSQL(cdsSeteoF, squery);
				}catch(Exception ex)
				{
					Funcion.LogSistema(cdsSeteoF, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), squery, Funcion.slogtipoerror);
				}

				//Guardar el documento
				DataView dvLineaCreditoGarantiaAvaluoDocCrear = new DataView();
				string squery2 = "";
				try
				{
					squery2 = string.Format("Exec LineaCreditoGarantiaAvaluoDocCrear '{0}', {1}, {2}, {3}", 
						txtDirectorioarchivo, 
						frmLineaCreditoCrear.idLineaCreditoTransaccion,
						idGarantiaLineaCreditoAvaluo,
						idGarantiaLineaCreditoAvaluoDoc
						);
					dvLineaCreditoGarantiaAvaluoDocCrear = Funcion.dvProcedimiento(cdsSeteoF, squery2);
				}
				catch(Exception ex)
				{
					Funcion.LogSistema(cdsSeteoF, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), squery2, Funcion.slogtipoerror);
				}
							  
				DataSet dsLineaCreditoGarantiaAvaluoDocCrear = new DataSet();
				dsLineaCreditoGarantiaAvaluoDocCrear = dvLineaCreditoGarantiaAvaluoDocCrear.Table.DataSet;
				bool bResultado = bool.Parse(dsLineaCreditoGarantiaAvaluoDocCrear.Tables[0].Rows[0]["bResultado"].ToString());
				string sMensaje = dsLineaCreditoGarantiaAvaluoDocCrear.Tables[0].Rows[0]["sMensaje"].ToString();

				if (idGarantiaLineaCreditoAvaluo != 0 && bResultado)
				{
					MessageBox.Show(string.Format("Proceso Finalizado Correctamente, se creo el registro exitosamente."), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.DialogResult = DialogResult.OK;
				}
			}
			#endregion Nuevo avalúo
			#region Eliminar avalúo
			if(bEliminar)
			{
				string squery3 = "";
				try
				{
					squery3 = String.Format("EXEC LineaCreditoGarantiaAvaluoEliminar {0}", idGarantiaLineaCreditoAvaluo);
					bool bResultadoEliminar = Funcion.bEscalarSQL(cdsSeteoF, String.Format("EXEC LineaCreditoGarantiaAvaluoEliminar {0}", idGarantiaLineaCreditoAvaluo), false);
					if(bResultadoEliminar)
					{
						MessageBox.Show(string.Format("Proceso Finalizado Correctamente, se elimino el registro exitosamente."), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.DialogResult = DialogResult.OK;
					}
				}
				catch(Exception ex)
				{
					Funcion.LogSistema(cdsSeteoF, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), squery3, Funcion.slogtipoerror);
				}
			}
			#endregion Eliminar avalúo
		}

		private void button4_Click(object sender, System.EventArgs e)
		{
					this.DialogResult = DialogResult.No;
		}

	}
}
