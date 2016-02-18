namespace ListViewEx
{
    using System;
    using System.Windows.Forms;

    public class SubItemEndEditingEventArgs : SubItemEventArgs
    {
        private bool _cancel;
        private string _text;

        public SubItemEndEditingEventArgs(ListViewItem item, int subItem, string display, bool cancel) : base(item, subItem)
        {
            this._text = string.Empty;
            this._cancel = true;
            this._text = display;
            this._cancel = cancel;
        }

        public bool Cancel
        {
            get
            {
                return this._cancel;
            }
            set
            {
                this._cancel = value;
            }
        }

        public string DisplayText
        {
            get
            {
                return this._text;
            }
            set
            {
                this._text = value;
            }
        }
    }
}
