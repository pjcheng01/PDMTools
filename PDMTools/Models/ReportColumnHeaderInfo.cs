using System.ComponentModel;

namespace PDMTools.Models
{
    /// <summary>結果 DataGrid 欄標題：標題文字 + 漏斗篩選按鈕綁定用。</summary>
    public sealed class ReportColumnHeaderInfo : INotifyPropertyChanged
    {
        private bool _isFiltered;

        public string Key { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;

        public bool IsFiltered
        {
            get => _isFiltered;
            set
            {
                if (_isFiltered == value)
                    return;
                _isFiltered = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsFiltered)));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
    }
}
