namespace PointSistemas
{
    partial class frmCre_ReporteRevisionDeCedulas
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle25 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle26 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle27 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle28 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle29 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle30 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle31 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle32 = new System.Windows.Forms.DataGridViewCellStyle();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtDesde = new System.Windows.Forms.MaskedTextBox();
            this.txtHasta = new System.Windows.Forms.MaskedTextBox();
            this.dgvDatos = new System.Windows.Forms.DataGridView();
            this.btnVer = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.lblContador = new System.Windows.Forms.Label();
            this.Local = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Consultas = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.SiAplica = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.NoAplica = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Llena = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.NoFacturado = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Facturado = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Efectividad = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnExportarExcel = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgvDatos)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 11);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(37, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Desde";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(154, 11);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(35, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Hasta";
            // 
            // txtDesde
            // 
            this.txtDesde.Location = new System.Drawing.Point(62, 8);
            this.txtDesde.Margin = new System.Windows.Forms.Padding(2, 4, 2, 4);
            this.txtDesde.Mask = "00/00/0000";
            this.txtDesde.Name = "txtDesde";
            this.txtDesde.Size = new System.Drawing.Size(68, 21);
            this.txtDesde.TabIndex = 0;
            this.txtDesde.ValidatingType = typeof(System.DateTime);
            this.txtDesde.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDesde_KeyDown);
            // 
            // txtHasta
            // 
            this.txtHasta.Location = new System.Drawing.Point(208, 8);
            this.txtHasta.Margin = new System.Windows.Forms.Padding(2, 4, 2, 4);
            this.txtHasta.Mask = "00/00/0000";
            this.txtHasta.Name = "txtHasta";
            this.txtHasta.Size = new System.Drawing.Size(68, 21);
            this.txtHasta.TabIndex = 1;
            this.txtHasta.ValidatingType = typeof(System.DateTime);
            this.txtHasta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtHasta_KeyDown);
            // 
            // dgvDatos
            // 
            this.dgvDatos.AllowUserToAddRows = false;
            this.dgvDatos.AllowUserToDeleteRows = false;
            this.dgvDatos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvDatos.BackgroundColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle25.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle25.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle25.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle25.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle25.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle25.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            this.dgvDatos.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle25;
            this.dgvDatos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvDatos.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Local,
            this.Consultas,
            this.SiAplica,
            this.NoAplica,
            this.Llena,
            this.NoFacturado,
            this.Facturado,
            this.Efectividad});
            this.dgvDatos.Location = new System.Drawing.Point(5, 55);
            this.dgvDatos.Margin = new System.Windows.Forms.Padding(2, 4, 2, 4);
            this.dgvDatos.Name = "dgvDatos";
            this.dgvDatos.ReadOnly = true;
            this.dgvDatos.RowHeadersWidth = 20;
            this.dgvDatos.Size = new System.Drawing.Size(746, 366);
            this.dgvDatos.TabIndex = 3;
            // 
            // btnVer
            // 
            this.btnVer.Location = new System.Drawing.Point(322, 6);
            this.btnVer.Margin = new System.Windows.Forms.Padding(2, 4, 2, 4);
            this.btnVer.Name = "btnVer";
            this.btnVer.Size = new System.Drawing.Size(74, 22);
            this.btnVer.TabIndex = 2;
            this.btnVer.Text = "Ver";
            this.btnVer.UseVisualStyleBackColor = true;
            this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Location = new System.Drawing.Point(2, 36);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(752, 8);
            this.groupBox1.TabIndex = 6;
            this.groupBox1.TabStop = false;
            // 
            // lblContador
            // 
            this.lblContador.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblContador.AutoSize = true;
            this.lblContador.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
            this.lblContador.Location = new System.Drawing.Point(6, 425);
            this.lblContador.Name = "lblContador";
            this.lblContador.Size = new System.Drawing.Size(0, 16);
            this.lblContador.TabIndex = 18;
            this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Local
            // 
            this.Local.DataPropertyName = "Local";
            this.Local.Frozen = true;
            this.Local.HeaderText = "Local";
            this.Local.Name = "Local";
            this.Local.ReadOnly = true;
            this.Local.Width = 160;
            // 
            // Consultas
            // 
            this.Consultas.DataPropertyName = "Consultas";
            dataGridViewCellStyle26.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle26.Format = "N0";
            dataGridViewCellStyle26.NullValue = "0";
            this.Consultas.DefaultCellStyle = dataGridViewCellStyle26;
            this.Consultas.HeaderText = "Consultas";
            this.Consultas.Name = "Consultas";
            this.Consultas.ReadOnly = true;
            this.Consultas.Width = 75;
            // 
            // SiAplica
            // 
            this.SiAplica.DataPropertyName = "SiAplica";
            dataGridViewCellStyle27.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle27.Format = "N0";
            dataGridViewCellStyle27.NullValue = "0";
            this.SiAplica.DefaultCellStyle = dataGridViewCellStyle27;
            this.SiAplica.HeaderText = "Si Aplica";
            this.SiAplica.Name = "SiAplica";
            this.SiAplica.ReadOnly = true;
            this.SiAplica.Width = 75;
            // 
            // NoAplica
            // 
            this.NoAplica.DataPropertyName = "NoAplica";
            dataGridViewCellStyle28.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle28.Format = "N0";
            this.NoAplica.DefaultCellStyle = dataGridViewCellStyle28;
            this.NoAplica.HeaderText = "No Aplica";
            this.NoAplica.Name = "NoAplica";
            this.NoAplica.ReadOnly = true;
            this.NoAplica.Width = 80;
            // 
            // Llena
            // 
            this.Llena.DataPropertyName = "Llena";
            dataGridViewCellStyle29.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle29.Format = "N0";
            dataGridViewCellStyle29.NullValue = "0";
            this.Llena.DefaultCellStyle = dataGridViewCellStyle29;
            this.Llena.HeaderText = "Sol Llena";
            this.Llena.Name = "Llena";
            this.Llena.ReadOnly = true;
            this.Llena.Width = 75;
            // 
            // NoFacturado
            // 
            this.NoFacturado.DataPropertyName = "NoFacturado";
            dataGridViewCellStyle30.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle30.Format = "N0";
            dataGridViewCellStyle30.NullValue = "0";
            this.NoFacturado.DefaultCellStyle = dataGridViewCellStyle30;
            this.NoFacturado.HeaderText = "No Facturado";
            this.NoFacturado.Name = "NoFacturado";
            this.NoFacturado.ReadOnly = true;
            this.NoFacturado.Width = 80;
            // 
            // Facturado
            // 
            this.Facturado.DataPropertyName = "Facturado";
            dataGridViewCellStyle31.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle31.Format = "N0";
            dataGridViewCellStyle31.NullValue = "0";
            this.Facturado.DefaultCellStyle = dataGridViewCellStyle31;
            this.Facturado.HeaderText = "Facturado";
            this.Facturado.Name = "Facturado";
            this.Facturado.ReadOnly = true;
            this.Facturado.Width = 75;
            // 
            // Efectividad
            // 
            this.Efectividad.DataPropertyName = "Efectividad";
            dataGridViewCellStyle32.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle32.Format = "N2";
            dataGridViewCellStyle32.NullValue = "0";
            this.Efectividad.DefaultCellStyle = dataGridViewCellStyle32;
            this.Efectividad.HeaderText = "% Efectividad";
            this.Efectividad.Name = "Efectividad";
            this.Efectividad.ReadOnly = true;
            this.Efectividad.Width = 80;
            // 
            // btnExportarExcel
            // 
            this.btnExportarExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExportarExcel.Location = new System.Drawing.Point(668, 8);
            this.btnExportarExcel.Name = "btnExportarExcel";
            this.btnExportarExcel.Size = new System.Drawing.Size(74, 23);
            this.btnExportarExcel.TabIndex = 21;
            this.btnExportarExcel.Text = "&Excel";
            this.btnExportarExcel.UseVisualStyleBackColor = true;
            this.btnExportarExcel.Click += new System.EventHandler(this.btnExportarExcel_Click);
            // 
            // frmCre_ReporteRevisionDeCedulas
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(754, 484);
            this.Controls.Add(this.btnExportarExcel);
            this.Controls.Add(this.lblContador);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btnVer);
            this.Controls.Add(this.dgvDatos);
            this.Controls.Add(this.txtHasta);
            this.Controls.Add(this.txtDesde);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.KeyPreview = true;
            this.Margin = new System.Windows.Forms.Padding(2, 4, 2, 4);
            this.MaximizeBox = false;
            this.Name = "frmCre_ReporteRevisionDeCedulas";
            this.Text = "Revision De Cedulas";
            this.Load += new System.EventHandler(this.frmCre_ReporteRevisionDeCedulas_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCre_ReporteRevisionDeCedulas_KeyDown);
            ((System.ComponentModel.ISupportInitialize)(this.dgvDatos)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.MaskedTextBox txtDesde;
        private System.Windows.Forms.MaskedTextBox txtHasta;
        private System.Windows.Forms.DataGridView dgvDatos;
        private System.Windows.Forms.Button btnVer;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label lblContador;
        private System.Windows.Forms.DataGridViewTextBoxColumn Local;
        private System.Windows.Forms.DataGridViewTextBoxColumn Consultas;
        private System.Windows.Forms.DataGridViewTextBoxColumn SiAplica;
        private System.Windows.Forms.DataGridViewTextBoxColumn NoAplica;
        private System.Windows.Forms.DataGridViewTextBoxColumn Llena;
        private System.Windows.Forms.DataGridViewTextBoxColumn NoFacturado;
        private System.Windows.Forms.DataGridViewTextBoxColumn Facturado;
        private System.Windows.Forms.DataGridViewTextBoxColumn Efectividad;
        private System.Windows.Forms.Button btnExportarExcel;
    }
}