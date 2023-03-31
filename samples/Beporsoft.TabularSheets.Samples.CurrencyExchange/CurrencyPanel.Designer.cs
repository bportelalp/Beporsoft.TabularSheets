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
            ((System.ComponentModel.ISupportInitialize)(this._dataCurrencies)).BeginInit();
            this.SuspendLayout();
            // 
            // _lblTitle
            // 
            this._lblTitle.AutoSize = true;
            this._lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this._lblTitle.Location = new System.Drawing.Point(12, 19);
            this._lblTitle.Name = "_lblTitle";
            this._lblTitle.Size = new System.Drawing.Size(203, 26);
            this._lblTitle.TabIndex = 0;
            this._lblTitle.Text = "Currency Exchange";
            // 
            // _lblOriginCurrency
            // 
            this._lblOriginCurrency.AutoSize = true;
            this._lblOriginCurrency.Location = new System.Drawing.Point(15, 89);
            this._lblOriginCurrency.Name = "_lblOriginCurrency";
            this._lblOriginCurrency.Size = new System.Drawing.Size(79, 13);
            this._lblOriginCurrency.TabIndex = 1;
            this._lblOriginCurrency.Text = "Origin Currency";
            // 
            // _dateTimeFrom
            // 
            this._dateTimeFrom.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this._dateTimeFrom.Location = new System.Drawing.Point(18, 372);
            this._dateTimeFrom.MinDate = new System.DateTime(2020, 1, 1, 0, 0, 0, 0);
            this._dateTimeFrom.Name = "_dateTimeFrom";
            this._dateTimeFrom.Size = new System.Drawing.Size(103, 20);
            this._dateTimeFrom.TabIndex = 2;
            // 
            // _dateTimeTo
            // 
            this._dateTimeTo.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this._dateTimeTo.Location = new System.Drawing.Point(133, 372);
            this._dateTimeTo.Name = "_dateTimeTo";
            this._dateTimeTo.Size = new System.Drawing.Size(115, 20);
            this._dateTimeTo.TabIndex = 3;
            // 
            // _comboOriginCurrency
            // 
            this._comboOriginCurrency.FormattingEnabled = true;
            this._comboOriginCurrency.Location = new System.Drawing.Point(17, 116);
            this._comboOriginCurrency.Name = "_comboOriginCurrency";
            this._comboOriginCurrency.Size = new System.Drawing.Size(164, 21);
            this._comboOriginCurrency.TabIndex = 4;
            // 
            // _dataCurrencies
            // 
            this._dataCurrencies.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this._dataCurrencies.Location = new System.Drawing.Point(409, 80);
            this._dataCurrencies.Name = "_dataCurrencies";
            this._dataCurrencies.Size = new System.Drawing.Size(310, 190);
            this._dataCurrencies.TabIndex = 5;
            // 
            // _listDestinationCurrencies
            // 
            this._listDestinationCurrencies.FormattingEnabled = true;
            this._listDestinationCurrencies.Location = new System.Drawing.Point(198, 116);
            this._listDestinationCurrencies.Name = "_listDestinationCurrencies";
            this._listDestinationCurrencies.Size = new System.Drawing.Size(173, 154);
            this._listDestinationCurrencies.TabIndex = 6;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(195, 89);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(112, 13);
            this.label1.TabIndex = 7;
            this.label1.Text = "Destination currencies";
            // 
            // _btnSearch
            // 
            this._btnSearch.Location = new System.Drawing.Point(18, 194);
            this._btnSearch.Name = "_btnSearch";
            this._btnSearch.Size = new System.Drawing.Size(163, 76);
            this._btnSearch.TabIndex = 8;
            this._btnSearch.Text = "Search";
            this._btnSearch.UseVisualStyleBackColor = true;
            this._btnSearch.Click += new System.EventHandler(this.OnSearch);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(13, 300);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(160, 26);
            this.label2.TabIndex = 9;
            this.label2.Text = "Export Historial";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(15, 348);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(30, 13);
            this.label3.TabIndex = 10;
            this.label3.Text = "From";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(130, 348);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(20, 13);
            this.label4.TabIndex = 11;
            this.label4.Text = "To";
            // 
            // _btnExport
            // 
            this._btnExport.Location = new System.Drawing.Point(274, 372);
            this._btnExport.Name = "_btnExport";
            this._btnExport.Size = new System.Drawing.Size(119, 20);
            this._btnExport.TabIndex = 12;
            this._btnExport.Text = "Export";
            this._btnExport.UseVisualStyleBackColor = true;
            this._btnExport.Click += new System.EventHandler(this.ExportSheetHistory);
            // 
            // btnSelectAll
            // 
            this.btnSelectAll.Location = new System.Drawing.Point(198, 276);
            this.btnSelectAll.Name = "btnSelectAll";
            this.btnSelectAll.Size = new System.Drawing.Size(81, 20);
            this.btnSelectAll.TabIndex = 13;
            this.btnSelectAll.Text = "Mark all";
            this.btnSelectAll.UseVisualStyleBackColor = true;
            this.btnSelectAll.Click += new System.EventHandler(this.BtnSelectAll);
            // 
            // btnUnselectAll
            // 
            this.btnUnselectAll.Location = new System.Drawing.Point(285, 276);
            this.btnUnselectAll.Name = "btnUnselectAll";
            this.btnUnselectAll.Size = new System.Drawing.Size(86, 20);
            this.btnUnselectAll.TabIndex = 14;
            this.btnUnselectAll.Text = "Unmark all";
            this.btnUnselectAll.UseVisualStyleBackColor = true;
            this.btnUnselectAll.Click += new System.EventHandler(this.BtnUnselectAll);
            // 
            // CurrencyPanel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(731, 450);
            this.Controls.Add(this.btnUnselectAll);
            this.Controls.Add(this.btnSelectAll);
            this.Controls.Add(this._btnExport);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this._btnSearch);
            this.Controls.Add(this.label1);
            this.Controls.Add(this._listDestinationCurrencies);
            this.Controls.Add(this._dataCurrencies);
            this.Controls.Add(this._comboOriginCurrency);
            this.Controls.Add(this._dateTimeTo);
            this.Controls.Add(this._dateTimeFrom);
            this.Controls.Add(this._lblOriginCurrency);
            this.Controls.Add(this._lblTitle);
            this.Name = "CurrencyPanel";
            this.Text = "Currency Exchange";
            ((System.ComponentModel.ISupportInitialize)(this._dataCurrencies)).EndInit();
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
    }
}

