using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de PlanCuentas.
	/// </summary>
	public class Seguridad : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsSeguridad;
		private C1.Data.C1DataView cdvSeguridad;
		private C1.Data.C1DataView cdvUsuario;
		private DevExpress.XtraTreeList.TreeList tlsSeguridad;
		private DevExpress.XtraTreeList.Columns.TreeListColumn tlcFuncion;
		private DevExpress.XtraTreeList.Columns.TreeListColumn tlcNivel;
		private System.Windows.Forms.Label label1;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private C1.Data.C1DataSet cdsUsuario;
		private DevExpress.XtraEditors.SimpleButton btSeguridad;
		private DevExpress.XtraTreeList.Columns.TreeListColumn tlcVisible;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnGrupo;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Seguridad()
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			this.cdvUsuario = new C1.Data.C1DataView();
			this.cdsUsuario = new C1.Data.C1DataSet();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			this.cdvSeguridad = new C1.Data.C1DataView();
			this.tlsSeguridad = new DevExpress.XtraTreeList.TreeList();
			this.tlcFuncion = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.tlcNivel = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.tlcVisible = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.label1 = new System.Windows.Forms.Label();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.btSeguridad = new DevExpress.XtraEditors.SimpleButton();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.spnGrupo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			((System.ComponentModel.ISupportInitialize)(this.cdvUsuario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsUsuario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSeguridad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.tlsSeguridad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnGrupo)).BeginInit();
			this.SuspendLayout();
			// 
			// cdvUsuario
			// 
			this.cdvUsuario.BindingContextCtrl = this;
			this.cdvUsuario.DataSet = this.cdsUsuario;
			this.cdvUsuario.Sort = "Nombre";
			this.cdvUsuario.TableName = "";
			this.cdvUsuario.TableViewName = "Usuario";
			// 
			// cdsUsuario
			// 
			this.cdsUsuario.BindingContextCtrl = this;
			this.cdsUsuario.DataLibrary = "LibEmpresa";
			this.cdsUsuario.DataLibraryUrl = "";
			this.cdsUsuario.DataSetDef = "dsUsuario";
			this.cdsUsuario.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsUsuario.Name = "cdsUsuario";
			this.cdsUsuario.SchemaClassName = "LibEmpresa.DataClass";
			this.cdsUsuario.SchemaDef = null;
			// 
			// cdsSeguridad
			// 
			this.cdsSeguridad.BindingContextCtrl = this;
			this.cdsSeguridad.DataLibrary = "LibContabilidad";
			this.cdsSeguridad.DataLibraryUrl = "";
			this.cdsSeguridad.DataSetDef = "dsSeguridad";
			this.cdsSeguridad.Locale = new System.Globalization.CultureInfo("");
			this.cdsSeguridad.Name = "cdsSeguridad";
			this.cdsSeguridad.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeguridad.SchemaDef = null;
			this.cdsSeguridad.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeguridad_BeforeFill);
			// 
			// cdvSeguridad
			// 
			this.cdvSeguridad.BindingContextCtrl = this;
			this.cdvSeguridad.DataSet = this.cdsSeguridad;
			this.cdvSeguridad.Sort = "Codigo";
			this.cdvSeguridad.TableName = "";
			this.cdvSeguridad.TableViewName = "Seguridad";
			// 
			// tlsSeguridad
			// 
			this.tlsSeguridad.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.tlsSeguridad.Appearance.Empty.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(243)), ((System.Byte)(243)), ((System.Byte)(243)));
			this.tlsSeguridad.Appearance.Empty.Options.UseBackColor = true;
			this.tlsSeguridad.Appearance.EvenRow.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(223)), ((System.Byte)(223)), ((System.Byte)(223)));
			this.tlsSeguridad.Appearance.EvenRow.BackColor2 = System.Drawing.Color.GhostWhite;
			this.tlsSeguridad.Appearance.EvenRow.Font = new System.Drawing.Font("Tahoma", 10F);
			this.tlsSeguridad.Appearance.EvenRow.ForeColor = System.Drawing.Color.Black;
			this.tlsSeguridad.Appearance.EvenRow.GradientMode = System.Drawing.Drawing2D.LinearGradientMode.ForwardDiagonal;
			this.tlsSeguridad.Appearance.EvenRow.Options.UseBackColor = true;
			this.tlsSeguridad.Appearance.EvenRow.Options.UseFont = true;
			this.tlsSeguridad.Appearance.EvenRow.Options.UseForeColor = true;
			this.tlsSeguridad.Appearance.FocusedCell.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(225)));
			this.tlsSeguridad.Appearance.FocusedCell.Font = new System.Drawing.Font("Tahoma", 10F);
			this.tlsSeguridad.Appearance.FocusedCell.ForeColor = System.Drawing.Color.Black;
			this.tlsSeguridad.Appearance.FocusedCell.Options.UseBackColor = true;
			this.tlsSeguridad.Appearance.FocusedCell.Options.UseFont = true;
			this.tlsSeguridad.Appearance.FocusedCell.Options.UseForeColor = true;
			this.tlsSeguridad.Appearance.FocusedRow.BackColor = System.Drawing.Color.Navy;
			this.tlsSeguridad.Appearance.FocusedRow.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(178)));
			this.tlsSeguridad.Appearance.FocusedRow.Font = new System.Drawing.Font("Tahoma", 10F);
			this.tlsSeguridad.Appearance.FocusedRow.ForeColor = System.Drawing.Color.White;
			this.tlsSeguridad.Appearance.FocusedRow.Options.UseBackColor = true;
			this.tlsSeguridad.Appearance.FocusedRow.Options.UseFont = true;
			this.tlsSeguridad.Appearance.FocusedRow.Options.UseForeColor = true;
			this.tlsSeguridad.Appearance.FooterPanel.BackColor = System.Drawing.Color.Silver;
			this.tlsSeguridad.Appearance.FooterPanel.BorderColor = System.Drawing.Color.Silver;
			this.tlsSeguridad.Appearance.FooterPanel.Font = new System.Drawing.Font("Tahoma", 10F);
			this.tlsSeguridad.Appearance.FooterPanel.ForeColor = System.Drawing.Color.Black;
			this.tlsSeguridad.Appearance.FooterPanel.Options.UseBackColor = true;
			this.tlsSeguridad.Appearance.FooterPanel.Options.UseBorderColor = true;
			this.tlsSeguridad.Appearance.FooterPanel.Options.UseFont = true;
			this.tlsSeguridad.Appearance.FooterPanel.Options.UseForeColor = true;
			this.tlsSeguridad.Appearance.GroupButton.BackColor = System.Drawing.Color.Silver;
			this.tlsSeguridad.Appearance.GroupButton.BorderColor = System.Drawing.Color.Silver;
			this.tlsSeguridad.Appearance.GroupButton.ForeColor = System.Drawing.Color.Black;
			this.tlsSeguridad.Appearance.GroupButton.Options.UseBackColor = true;
			this.tlsSeguridad.Appearance.GroupButton.Options.UseBorderColor = true;
			this.tlsSeguridad.Appearance.GroupButton.Options.UseForeColor = true;
			this.tlsSeguridad.Appearance.GroupFooter.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(202)), ((System.Byte)(202)), ((System.Byte)(202)));
			this.tlsSeguridad.Appearance.GroupFooter.BorderColor = System.Drawing.Color.FromArgb(((System.Byte)(202)), ((System.Byte)(202)), ((System.Byte)(202)));
			this.tlsSeguridad.Appearance.GroupFooter.ForeColor = System.Drawing.Color.Black;
			this.tlsSeguridad.Appearance.GroupFooter.Options.UseBackColor = true;
			this.tlsSeguridad.Appearance.GroupFooter.Options.UseBorderColor = true;
			this.tlsSeguridad.Appearance.GroupFooter.Options.UseForeColor = true;
			this.tlsSeguridad.Appearance.HeaderPanel.BackColor = System.Drawing.Color.Silver;
			this.tlsSeguridad.Appearance.HeaderPanel.BorderColor = System.Drawing.Color.Silver;
			this.tlsSeguridad.Appearance.HeaderPanel.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold);
			this.tlsSeguridad.Appearance.HeaderPanel.ForeColor = System.Drawing.Color.Black;
			this.tlsSeguridad.Appearance.HeaderPanel.Options.UseBackColor = true;
			this.tlsSeguridad.Appearance.HeaderPanel.Options.UseBorderColor = true;
			this.tlsSeguridad.Appearance.HeaderPanel.Options.UseFont = true;
			this.tlsSeguridad.Appearance.HeaderPanel.Options.UseForeColor = true;
			this.tlsSeguridad.Appearance.HideSelectionRow.BackColor = System.Drawing.Color.Gray;
			this.tlsSeguridad.Appearance.HideSelectionRow.Font = new System.Drawing.Font("Tahoma", 10F);
			this.tlsSeguridad.Appearance.HideSelectionRow.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200)));
			this.tlsSeguridad.Appearance.HideSelectionRow.Options.UseBackColor = true;
			this.tlsSeguridad.Appearance.HideSelectionRow.Options.UseFont = true;
			this.tlsSeguridad.Appearance.HideSelectionRow.Options.UseForeColor = true;
			this.tlsSeguridad.Appearance.HorzLine.BackColor = System.Drawing.Color.Silver;
			this.tlsSeguridad.Appearance.HorzLine.Options.UseBackColor = true;
			this.tlsSeguridad.Appearance.OddRow.BackColor = System.Drawing.Color.White;
			this.tlsSeguridad.Appearance.OddRow.BackColor2 = System.Drawing.Color.White;
			this.tlsSeguridad.Appearance.OddRow.Font = new System.Drawing.Font("Tahoma", 10F);
			this.tlsSeguridad.Appearance.OddRow.ForeColor = System.Drawing.Color.Black;
			this.tlsSeguridad.Appearance.OddRow.GradientMode = System.Drawing.Drawing2D.LinearGradientMode.BackwardDiagonal;
			this.tlsSeguridad.Appearance.OddRow.Options.UseBackColor = true;
			this.tlsSeguridad.Appearance.OddRow.Options.UseFont = true;
			this.tlsSeguridad.Appearance.OddRow.Options.UseForeColor = true;
			this.tlsSeguridad.Appearance.Preview.BackColor = System.Drawing.Color.White;
			this.tlsSeguridad.Appearance.Preview.Font = new System.Drawing.Font("Tahoma", 10F);
			this.tlsSeguridad.Appearance.Preview.ForeColor = System.Drawing.Color.Navy;
			this.tlsSeguridad.Appearance.Preview.Options.UseBackColor = true;
			this.tlsSeguridad.Appearance.Preview.Options.UseFont = true;
			this.tlsSeguridad.Appearance.Preview.Options.UseForeColor = true;
			this.tlsSeguridad.Appearance.Row.BackColor = System.Drawing.Color.White;
			this.tlsSeguridad.Appearance.Row.Font = new System.Drawing.Font("Tahoma", 10F);
			this.tlsSeguridad.Appearance.Row.ForeColor = System.Drawing.Color.Black;
			this.tlsSeguridad.Appearance.Row.Options.UseBackColor = true;
			this.tlsSeguridad.Appearance.Row.Options.UseFont = true;
			this.tlsSeguridad.Appearance.Row.Options.UseForeColor = true;
			this.tlsSeguridad.Appearance.SelectedRow.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(10)), ((System.Byte)(10)), ((System.Byte)(138)));
			this.tlsSeguridad.Appearance.SelectedRow.Font = new System.Drawing.Font("Tahoma", 10F);
			this.tlsSeguridad.Appearance.SelectedRow.ForeColor = System.Drawing.Color.White;
			this.tlsSeguridad.Appearance.SelectedRow.Options.UseBackColor = true;
			this.tlsSeguridad.Appearance.SelectedRow.Options.UseFont = true;
			this.tlsSeguridad.Appearance.SelectedRow.Options.UseForeColor = true;
			this.tlsSeguridad.Appearance.TreeLine.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(58)), ((System.Byte)(58)), ((System.Byte)(58)));
			this.tlsSeguridad.Appearance.TreeLine.Options.UseBackColor = true;
			this.tlsSeguridad.Appearance.VertLine.BackColor = System.Drawing.Color.Silver;
			this.tlsSeguridad.Appearance.VertLine.Options.UseBackColor = true;
			this.tlsSeguridad.Columns.AddRange(new DevExpress.XtraTreeList.Columns.TreeListColumn[] {
																																																this.tlcFuncion,
																																																this.tlcNivel,
																																																this.tlcVisible});
			this.tlsSeguridad.DataSource = this.cdvSeguridad;
			this.tlsSeguridad.KeyFieldName = "id_seg";
			this.tlsSeguridad.Location = new System.Drawing.Point(176, 0);
			this.tlsSeguridad.Name = "tlsSeguridad";
			this.tlsSeguridad.OptionsView.EnableAppearanceEvenRow = true;
			this.tlsSeguridad.OptionsView.EnableAppearanceOddRow = true;
			this.tlsSeguridad.ParentFieldName = "padre";
			this.tlsSeguridad.Size = new System.Drawing.Size(496, 360);
			this.tlsSeguridad.Styles.AddReplace("PressedColumn", new DevExpress.Utils.ViewStyle("PressedColumn", "TreeList", new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold), "HeaderPanel", ((DevExpress.Utils.StyleOptions)(((DevExpress.Utils.StyleOptions.StyleEnabled | DevExpress.Utils.StyleOptions.UseBackColor) 
				| DevExpress.Utils.StyleOptions.UseForeColor))), true, false, false, DevExpress.Utils.HorzAlignment.Near, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(128))), System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255)))));
			this.tlsSeguridad.TabIndex = 3;
			this.tlsSeguridad.FocusedNodeChanged += new DevExpress.XtraTreeList.FocusedNodeChangedEventHandler(this.tlsSeguridad_FocusedNodeChanged);
			// 
			// tlcFuncion
			// 
			this.tlcFuncion.Caption = "Función";
			this.tlcFuncion.FieldName = "nombre";
			this.tlcFuncion.Name = "tlcFuncion";
			this.tlcFuncion.OptionsColumn.ReadOnly = true;
			this.tlcFuncion.VisibleIndex = 0;
			this.tlcFuncion.Width = 250;
			// 
			// tlcNivel
			// 
			this.tlcNivel.Caption = "Acceso";
			this.tlcNivel.FieldName = "n_1";
			this.tlcNivel.Name = "tlcNivel";
			this.tlcNivel.VisibleIndex = 1;
			this.tlcNivel.Width = 20;
			// 
			// tlcVisible
			// 
			this.tlcVisible.Caption = "Visible";
			this.tlcVisible.FieldName = "Vn_1";
			this.tlcVisible.Name = "tlcVisible";
			this.tlcVisible.OptionsColumn.AllowEdit = false;
			this.tlcVisible.VisibleIndex = 2;
			this.tlcVisible.Width = 30;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(39, 16);
			this.label1.TabIndex = 4;
			this.label1.Text = "Grupo:";
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.Location = new System.Drawing.Point(8, 336);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 25);
			this.btAceptar.TabIndex = 5;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btCancelar.Location = new System.Drawing.Point(96, 336);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 25);
			this.btCancelar.TabIndex = 6;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btSeguridad
			// 
			this.btSeguridad.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btSeguridad.Location = new System.Drawing.Point(8, 304);
			this.btSeguridad.Name = "btSeguridad";
			this.btSeguridad.Size = new System.Drawing.Size(88, 24);
			this.btSeguridad.TabIndex = 7;
			this.btSeguridad.Text = "Con&figuracion";
			this.btSeguridad.Click += new System.EventHandler(this.btSeguridad_Click);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance2;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			this.ultraGrid1.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.AllRowsInBand;
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
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.ultraGrid1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.ultraGrid1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.ultraGrid1.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.ultraGrid1.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 32);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(160, 264);
			this.ultraGrid1.TabIndex = 8;
			// 
			// spnGrupo
			// 
			this.spnGrupo.Location = new System.Drawing.Point(64, 8);
			this.spnGrupo.MaskInput = "nn";
			this.spnGrupo.MaxValue = 36;
			this.spnGrupo.MinValue = 1;
			this.spnGrupo.Name = "spnGrupo";
			this.spnGrupo.PromptChar = ' ';
			this.spnGrupo.Size = new System.Drawing.Size(104, 22);
			this.spnGrupo.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnGrupo.TabIndex = 9;
			this.spnGrupo.ValueChanged += new System.EventHandler(this.spnGrupo_TextChanged);
			// 
			// Seguridad
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(679, 366);
			this.Controls.Add(this.spnGrupo);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.btSeguridad);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.tlsSeguridad);
			this.Name = "Seguridad";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Seguridad";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.Seguridad_Closing);
			this.Load += new System.EventHandler(this.Seguridad_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdvUsuario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsUsuario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSeguridad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.tlsSeguridad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnGrupo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			this.cdsSeguridad.Update();
			Funcion.EjecutaSQL(cdsSeguridad, "Exec SeguridadVisible");
			this.Close();
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void cdsSeguridad_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeguridad.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void spnGrupo_TextChanged(object sender, System.EventArgs e)
		{
			//this.cdvUsuario.RowFilter = "idGrupo = " + this.spnGrupo.Value.ToString().Trim();
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeguridad, string.Format("Select Nombre From Usuario Where idGrupo = {0} And Activo = 1", (int)this.spnGrupo.Value));
			this.tlcNivel.FieldName = "n_" + this.spnGrupo.Value.ToString().Trim();
			this.tlcVisible.FieldName = "Vn_" + this.spnGrupo.Value.ToString().Trim();			
		}

		private void Seguridad_Load(object sender, System.EventArgs e)
		{			
			string stAudita = "Exec AuditaCrear 53, 6, 'Seguridad'";
			Funcion.EjecutaSQL(cdsSeguridad, stAudita, true);

			Cursor = Cursors.Default;
			spnGrupo.Value = MenuLatinium.iNivel;
			tlcNivel.FieldName = "n_" + MenuLatinium.iNivel.ToString().Trim();
			this.tlcVisible.FieldName = "Vn_" + MenuLatinium.iNivel.ToString().Trim();

			tlcVisible.OptionsColumn.AllowEdit = true;
			tlcNivel.OptionsColumn.AllowEdit = true;

			//tlcVisible.Width = 0;
			//if (MenuLatinium.stUsuario != "prg") btSeguridad.Width = 0;
			//if (!Funcion.Permiso("202", cdsSeguridad)) spnGrupo.Enabled = false;			
		}

		private void btSeguridad_Click(object sender, System.EventArgs e)
		{			
			cdvSeguridad.RowFilter = "";
		}

		private void Seguridad_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Seguridad'";
			Funcion.EjecutaSQL(cdsSeguridad, stAudita, true);
		}

		private void tlsSeguridad_FocusedNodeChanged(object sender, DevExpress.XtraTreeList.FocusedNodeChangedEventArgs e)
		{
		
		}

	}
}
