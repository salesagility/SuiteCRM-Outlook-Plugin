namespace ListViewEx
{
    using System;
    using System.Windows.Forms;

    public class SubItemEventArgs : EventArgs
    {
        private ListViewItem _item = null;
        private int _subItemIndex = -1;

        public SubItemEventArgs(ListViewItem item, int subItem)
        {
            this._subItemIndex = subItem;
            this._item = item;
        }

        public ListViewItem Item
        {
            get
            {
                return this._item;
            }
        }

        public int SubItem
        {
            get
            {
                return this._subItemIndex;
            }
        }
    }
}
