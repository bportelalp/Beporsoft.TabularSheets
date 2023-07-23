namespace Beporsoft.TabularSheets.Samples.CurrencyExchange
{
    partial class CurrencyPanel
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de Windows Forms

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this._lblTitle = new System.Windows.Forms.Label();
            this._lblOriginCurrency = new System.Windows.Forms.Label();
            this._dateTimeFrom = new System.Windows.Forms.DateTimePicker();
            this._dateTimeTo = new System.Windows.Forms.DateTimePicker();
            this._comboOriginCurrency = new System.Windows.Forms.ComboBox();
            this._dataCurrencies = new System.Windows.Forms.DataGridView();
            this._listDestinationCurrencies = new System.Windows.Forms.CheckedListBox();
            this.label1 = new System.Windows.Forms.Label();
            this._btnSearch = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this._btnExport = new System.Windows.Forms.Button();
            this.btnSelectAll = new System.Windows.Forms.Button();
            this.btnUnselectAll = new System.Windows.Forms.Button();
            this._saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.label5 = new System.Windows.Forms.Label();
            this.lblHeaderColor = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.btnConfigureHeaderFont = new System.Windows.Forms.Button();
            this.dialogFontHeader = new System.Windows.Forms.FontDialog();
            this.comboHeaderFill = new System.Windows.Forms.ComboBox();
            ((System.ComponentModel.ISupportInitialize)(this._dataCurrencies)).BeginInit();
            this.tableLayoutPanel1.SuspendLayout();
            this.tableLayoutPanel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // _lblTitle
            // 
            this._lblTitle.AutoSize = true;
            this._lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this._lblTitle.Location = new System.Drawing.Point(16, 26);
            this._lblTitle.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this._lblTitle.Name = "_lblTitle";
            this._lblTitle.Size = new System.Drawing.Size(203, 26);
            this._lblTitle.TabIndex = 0;
            this._lblTitle.Text = "Currency Exchange";
            // 
            // _lblOriginCurrency
            // 
            this._lblOriginCurrency.AutoSize = true;
            this._lblOriginCurrency.Font = new System.Drawing.Font("Calibri", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this._lblOriginCurrency.Location = new System.Drawing.Point(20, 123);
            this._lblOriginCurrency.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this._lblOriginCurrency.Name = "_lblOriginCurrency";
            this._lblOriginCurrency.Size = new System.Drawing.Size(104, 18);
            this._lblOriginCurrency.TabIndex = 1;
            this._lblOriginCurrency.Text = "Origin Currency";
            // 
            // _dateTimeFrom
            // 
            this._dateTimeFrom.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this._dateTimeFrom.Location = new System.Drawing.Point(65, 4);
            this._dateTimeFrom.Margin = new System.Windows.Forms.Padding(4);
            this._dateTimeFrom.MinDate = new System.DateTime(2020, 1, 1, 0, 0, 0, 0);
            this._dateTimeFrom.Name = "_dateTimeFrom";
            this._dateTimeFrom.Size = new System.Drawing.Size(119, 25);
            this._dateTimeFrom.TabIndex = 2;
            // 
            // _dateTimeTo
            // 
            this._dateTimeTo.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this._dateTimeTo.Location = new System.Drawing.Point(65, 41);
            this._dateTimeTo.Margin = new System.Windows.Forms.Padding(4);
            this._dateTimeTo.Name = "_dateTimeTo";
            this._dateTimeTo.Size = new System.Drawing.Size(119, 25);
            this._dateTimeTo.TabIndex = 3;
            // 
            // _comboOriginCurrency
            // 
            this._comboOriginCurrency.FormattingEnabled = true;
            this._comboOriginCurrency.Location = new System.Drawing.Point(23, 161);
            this._comboOriginCurrency.Margin = new System.Windows.Forms.Padding(4);
            this._comboOriginCurrency.Name = "_comboOriginCurrency";
            this._comboOriginCurrency.Size = new System.Drawing.Size(217, 26);
            this._comboOriginCurrency.TabIndex = 4;
            // 
            // _dataCurrencies
            // 
            this._dataCurrencies.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this._dataCurrencies.Location = new System.Drawing.Point(545, 111);
            this._dataCurrencies.Margin = new System.Windows.Forms.Padding(4);
            this._dataCurrencies.Name = "_dataCurrencies";
            this._dataCurrencies.Size = new System.Drawing.Size(413, 263);
            this._dataCurrencies.TabIndex = 5;
            // 
            // _listDestinationCurrencies
            // 
            this._listDestinationCurrencies.Font = new System.Drawing.Font("Calibri", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this._listDestinationCurrencies.FormattingEnabled = true;
            this._listDestinationCurrencies.Location = new System.Drawing.Point(264, 161);
            this._listDestinationCurrencies.Margin = new System.Windows.Forms.Padding(4);
            this._listDestinationCurrencies.Name = "_listDestinationCurrencies";
            this._listDestinationCurrencies.Size = new System.Drawing.Size(229, 212);
            this._listDestinationCurrencies.TabIndex = 6;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Calibri", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(260, 123);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(145, 18);
            this.label1.TabIndex = 7;
            this.label1.Text = "Destination currencies";
            // 
            // _btnSearch
            // 
            this._btnSearch.Location = new System.Drawing.Point(24, 269);
            this._btnSearch.Margin = new System.Windows.Forms.Padding(4);
            this._btnSearch.Name = "_btnSearch";
            this._btnSearch.Size = new System.Drawing.Size(217, 105);
            this._btnSearch.TabIndex = 8;
            this._btnSearch.Text = "Search";
            this._btnSearch.UseVisualStyleBackColor = true;
            this._btnSearch.Click += new System.EventHandler(this.OnClickSearch);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(17, 415);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(160, 26);
            this.label2.TabIndex = 9;
            this.label2.Text = "Export Historial";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label3.Font = new System.Drawing.Font("Calibri", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(4, 0);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 37);
            this.label3.TabIndex = 10;
            this.label3.Text = "From";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label4.Font = new System.Drawing.Font("Calibri", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(4, 37);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(53, 35);
            this.label4.TabIndex = 11;
            this.label4.Text = "To";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // _btnExport
            // 
            this._btnExport.Location = new System.Drawing.Point(229, 458);
            this._btnExport.Margin = new System.Windows.Forms.Padding(4);
            this._btnExport.Name = "_btnExport";
            this._btnExport.Size = new System.Drawing.Size(159, 28);
            this._btnExport.TabIndex = 12;
            this._btnExport.Text = "Export";
            this._btnExport.UseVisualStyleBackColor = true;
            this._btnExport.Click += new System.EventHandler(this.OnClickExportSheetHistory);
            // 
            // btnSelectAll
            // 
            this.btnSelectAll.Location = new System.Drawing.Point(264, 382);
            this.btnSelectAll.Margin = new System.Windows.Forms.Padding(4);
            this.btnSelectAll.Name = "btnSelectAll";
            this.btnSelectAll.Size = new System.Drawing.Size(108, 28);
            this.btnSelectAll.TabIndex = 13;
            this.btnSelectAll.Text = "Mark all";
            this.btnSelectAll.UseVisualStyleBackColor = true;
            this.btnSelectAll.Click += new System.EventHandler(this.OnClickSelectAll);
            // 
            // btnUnselectAll
            // 
            this.btnUnselectAll.Location = new System.Drawing.Point(380, 382);
            this.btnUnselectAll.Margin = new System.Windows.Forms.Padding(4);
            this.btnUnselectAll.Name = "btnUnselectAll";
            this.btnUnselectAll.Size = new System.Drawing.Size(115, 28);
            this.btnUnselectAll.TabIndex = 14;
            this.btnUnselectAll.Text = "Unmark all";
            this.btnUnselectAll.UseVisualStyleBackColor = true;
            this.btnUnselectAll.Click += new System.EventHandler(this.OnClickUnselectAll);
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 32.80423F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 67.19577F));
            this.tableLayoutPanel1.Controls.Add(this._dateTimeFrom, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this._dateTimeTo, 1, 1);
            this.tableLayoutPanel1.Controls.Add(this.label3, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.label4, 0, 1);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(21, 458);
            this.tableLayoutPanel1.Margin = new System.Windows.Forms.Padding(4);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 51.76471F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 48.23529F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(189, 72);
            this.tableLayoutPanel1.TabIndex = 15;
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.ColumnCount = 2;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 39.77273F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 60.22727F));
            this.tableLayoutPanel2.Controls.Add(this.label5, 0, 1);
            this.tableLayoutPanel2.Controls.Add(this.lblHeaderColor, 0, 0);
            this.tableLayoutPanel2.Controls.Add(this.label6, 0, 2);
            this.tableLayoutPanel2.Controls.Add(this.btnConfigureHeaderFont, 1, 2);
            this.tableLayoutPanel2.Controls.Add(this.comboHeaderFill, 1, 0);
            this.tableLayoutPanel2.Location = new System.Drawing.Point(508, 430);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 3;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 39F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(203, 121);
            this.tableLayoutPanel2.TabIndex = 16;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label5.Location = new System.Drawing.Point(3, 41);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(74, 41);
            this.label5.TabIndex = 1;
            this.label5.Text = "Border";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblHeaderColor
            // 
            this.lblHeaderColor.AutoSize = true;
            this.lblHeaderColor.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblHeaderColor.Location = new System.Drawing.Point(3, 0);
            this.lblHeaderColor.Name = "lblHeaderColor";
            this.lblHeaderColor.Size = new System.Drawing.Size(74, 41);
            this.lblHeaderColor.TabIndex = 0;
            this.lblHeaderColor.Text = "Fill color";
            this.lblHeaderColor.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label6.Location = new System.Drawing.Point(3, 82);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(74, 39);
            this.label6.TabIndex = 2;
            this.label6.Text = "Font";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // btnConfigureHeaderFont
            // 
            this.btnConfigureHeaderFont.Location = new System.Drawing.Point(83, 85);
            this.btnConfigureHeaderFont.Name = "btnConfigureHeaderFont";
            this.btnConfigureHeaderFont.Size = new System.Drawing.Size(97, 33);
            this.btnConfigureHeaderFont.TabIndex = 3;
            this.btnConfigureHeaderFont.Text = "Font";
            this.btnConfigureHeaderFont.UseVisualStyleBackColor = true;
            this.btnConfigureHeaderFont.Click += new System.EventHandler(this.OnClickConfigureHeaderFont);
            // 
            // dialogFontHeader
            // 
            this.dialogFontHeader.AllowScriptChange = false;
            this.dialogFontHeader.AllowSimulations = false;
            this.dialogFontHeader.ShowApply = true;
            this.dialogFontHeader.ShowColor = true;
            // 
            // comboHeaderFill
            // 
            this.comboHeaderFill.FormattingEnabled = true;
            this.comboHeaderFill.Location = new System.Drawing.Point(83, 3);
            this.comboHeaderFill.Name = "comboHeaderFill";
            this.comboHeaderFill.Size = new System.Drawing.Size(117, 26);
            this.comboHeaderFill.TabIndex = 4;
            // 
            // CurrencyPanel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(975, 623);
            this.Controls.Add(this.tableLayoutPanel2);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Controls.Add(this.btnUnselectAll);
            this.Controls.Add(this.btnSelectAll);
            this.Controls.Add(this._btnExport);
            this.Controls.Add(this.label2);
            this.Controls.Add(this._btnSearch);
            this.Controls.Add(this.label1);
            this.Controls.Add(this._listDestinationCurrencies);
            this.Controls.Add(this._dataCurrencies);
            this.Controls.Add(this._comboOriginCurrency);
            this.Controls.Add(this._lblOriginCurrency);
            this.Controls.Add(this._lblTitle);
            this.Font = new System.Drawing.Font("Calibri", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "CurrencyPanel";
            this.Text = "Currency Exchange";
            ((System.ComponentModel.ISupportInitialize)(this._dataCurrencies)).EndInit();
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.tableLayoutPanel2.ResumeLayout(false);
            this.tableLayoutPanel2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label _lblTitle;
        private System.Windows.Forms.Label _lblOriginCurrency;
        private System.Windows.Forms.DateTimePicker _dateTimeFrom;
        private System.Windows.Forms.DateTimePicker _dateTimeTo;
        private System.Windows.Forms.ComboBox _comboOriginCurrency;
        private System.Windows.Forms.DataGridView _dataCurrencies;
        private System.Windows.Forms.CheckedListBox _listDestinationCurrencies;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button _btnSearch;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button _btnExport;
        private System.Windows.Forms.Button btnSelectAll;
        private System.Windows.Forms.Button btnUnselectAll;
        private System.Windows.Forms.SaveFileDialog _saveFileDialog;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label lblHeaderColor;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button btnConfigureHeaderFont;
        private System.Windows.Forms.FontDialog dialogFontHeader;
        private System.Windows.Forms.ComboBox comboHeaderFill;
    }
}

