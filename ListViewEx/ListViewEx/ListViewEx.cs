namespace ListViewEx
{
    using System;
    using System.ComponentModel;
    using System.Drawing;
    using System.Runtime.CompilerServices;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;

    public class ListViewEx : ListView
    {
        private bool _doubleClickActivation = false;
        private Control _editingControl;
        private ListViewItem _editItem;
        private int _editSubItem;
        private Container components = null;
        private const int HDN_BEGINDRAG = -310;
        private const int HDN_FIRST = -300;
        private const int HDN_ITEMCHANGINGA = -300;
        private const int HDN_ITEMCHANGINGW = -320;
        private const int LVM_FIRST = 0x1000;
        private const int LVM_GETCOLUMNORDERARRAY = 0x103b;
        private const int WM_HSCROLL = 0x114;
        private const int WM_NOTIFY = 0x4e;
        private const int WM_SIZE = 5;
        private const int WM_VSCROLL = 0x115;

        public event SubItemEventHandler SubItemBeginEditing;

        public event SubItemEventHandler SubItemClicked;

        public event SubItemEndEditingEventHandler SubItemEndEditing;

        public ListViewEx()
        {
            this.InitializeComponent();
            base.FullRowSelect = true;
            base.View = View.Details;
            base.AllowColumnReorder = true;
        }

        private void _editControl_KeyPress(object sender, KeyPressEventArgs e)
        {
            switch (e.KeyChar)
            {
                case '\r':
                    this.EndEditing(true);
                    break;

                case '\x001b':
                    this.EndEditing(false);
                    break;
            }
        }

        private void _editControl_Leave(object sender, EventArgs e)
        {
            this.EndEditing(true);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && (this.components != null))
            {
                this.components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void EditSubitemAt(Point p)
        {
            ListViewItem item;
            int subItem = this.GetSubItemAt(p.X, p.Y, out item);
            if (subItem >= 0)
            {
                this.OnSubItemClicked(new SubItemEventArgs(item, subItem));
            }
        }

        public void EndEditing(bool AcceptChanges)
        {
            if (this._editingControl != null)
            {
                SubItemEndEditingEventArgs e = new SubItemEndEditingEventArgs(this._editItem, this._editSubItem, AcceptChanges ? this._editingControl.Text : this._editItem.SubItems[this._editSubItem].Text, !AcceptChanges);
                this.OnSubItemEndEditing(e);
                this._editItem.SubItems[this._editSubItem].Text = e.DisplayText;
                this._editingControl.Leave -= new EventHandler(this._editControl_Leave);
                this._editingControl.KeyPress -= new KeyPressEventHandler(this._editControl_KeyPress);
                this._editingControl.Visible = false;
                this._editingControl = null;
                this._editItem = null;
                this._editSubItem = -1;
            }
        }

        public int[] GetColumnOrder()
        {
            IntPtr lPar = Marshal.AllocHGlobal((int) (Marshal.SizeOf(typeof(int)) * base.Columns.Count));
            if (SendMessage(base.Handle, 0x103b, new IntPtr(base.Columns.Count), lPar).ToInt32() == 0)
            {
                Marshal.FreeHGlobal(lPar);
                return null;
            }
            int[] destination = new int[base.Columns.Count];
            Marshal.Copy(lPar, destination, 0, base.Columns.Count);
            Marshal.FreeHGlobal(lPar);
            return destination;
        }

        public int GetSubItemAt(int x, int y, out ListViewItem item)
        {
            item = base.GetItemAt(x, y);
            if (item != null)
            {
                int[] columnOrder = this.GetColumnOrder();
                int left = item.GetBounds(ItemBoundsPortion.Entire).Left;
                for (int i = 0; i < columnOrder.Length; i++)
                {
                    ColumnHeader header = base.Columns[columnOrder[i]];
                    if (x < (left + header.Width))
                    {
                        return header.Index;
                    }
                    left += header.Width;
                }
            }
            return -1;
        }

        public Rectangle GetSubItemBounds(ListViewItem Item, int SubItem)
        {
            int[] columnOrder = this.GetColumnOrder();
            if (SubItem >= columnOrder.Length)
            {
                throw new IndexOutOfRangeException("SubItem " + SubItem + " out of range");
            }
            if (Item == null)
            {
                throw new ArgumentNullException("Item");
            }
            Rectangle bounds = Item.GetBounds(ItemBoundsPortion.Entire);
            int left = bounds.Left;
            int index = 0;
            while (index < columnOrder.Length)
            {
                ColumnHeader header = base.Columns[columnOrder[index]];
                if (header.Index == SubItem)
                {
                    break;
                }
                left += header.Width;
                index++;
            }
            return new Rectangle(left, bounds.Top, base.Columns[columnOrder[index]].Width, bounds.Height);
        }

        private void InitializeComponent()
        {
            this.components = new Container();
        }

        protected override void OnDoubleClick(EventArgs e)
        {
            base.OnDoubleClick(e);
            if (this.DoubleClickActivation)
            {
                Point p = base.PointToClient(Cursor.Position);
                this.EditSubitemAt(p);
            }
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            if (!this.DoubleClickActivation)
            {
                this.EditSubitemAt(new Point(e.X, e.Y));
            }
        }

        protected void OnSubItemBeginEditing(SubItemEventArgs e)
        {
            if (this.SubItemBeginEditing != null)
            {
                this.SubItemBeginEditing(this, e);
            }
        }

        protected void OnSubItemClicked(SubItemEventArgs e)
        {
            if (this.SubItemClicked != null)
            {
                this.SubItemClicked(this, e);
            }
        }

        protected void OnSubItemEndEditing(SubItemEndEditingEventArgs e)
        {
            if (this.SubItemEndEditing != null)
            {
                this.SubItemEndEditing(this, e);
            }
        }

        [DllImport("user32.dll", CharSet=CharSet.Ansi)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, int len, ref int[] order);
        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wPar, IntPtr lPar);
        public void StartEditing(Control c, ListViewItem Item, int SubItem)
        {
            this.OnSubItemBeginEditing(new SubItemEventArgs(Item, SubItem));
            Rectangle subItemBounds = this.GetSubItemBounds(Item, SubItem);
            if (subItemBounds.X < 0)
            {
                subItemBounds.Width += subItemBounds.X;
                subItemBounds.X = 0;
            }
            if ((subItemBounds.X + subItemBounds.Width) > base.Width)
            {
                subItemBounds.Width = base.Width - subItemBounds.Left;
            }
            subItemBounds.Offset(base.Left, base.Top);
            Point p = new Point(0, 0);
            Point point2 = base.Parent.PointToScreen(p);
            Point point3 = c.Parent.PointToScreen(p);
            subItemBounds.Offset(point2.X - point3.X, point2.Y - point3.Y);
            c.Bounds = subItemBounds;
            c.Text = Item.SubItems[SubItem].Text;
            c.Visible = true;
            c.BringToFront();
            c.Focus();
            this._editingControl = c;
            this._editingControl.Leave += new EventHandler(this._editControl_Leave);
            this._editingControl.KeyPress += new KeyPressEventHandler(this._editControl_KeyPress);
            this._editItem = Item;
            this._editSubItem = SubItem;
        }

        protected override void WndProc(ref Message msg)
        {
            switch (msg.Msg)
            {
                case 0x114:
                case 0x115:
                case 5:
                    this.EndEditing(false);
                    break;

                case 0x4e:
                {
                    NMHDR nmhdr = (NMHDR) Marshal.PtrToStructure(msg.LParam, typeof(NMHDR));
                    if (((nmhdr.code == -310) || (nmhdr.code == -300)) || (nmhdr.code == -320))
                    {
                        this.EndEditing(false);
                    }
                    break;
                }
            }
            base.WndProc(ref msg);
        }

        public bool DoubleClickActivation
        {
            get
            {
                return this._doubleClickActivation;
            }
            set
            {
                this._doubleClickActivation = value;
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct NMHDR
        {
            public IntPtr hwndFrom;
            public int idFrom;
            public int code;
        }
    }
}
