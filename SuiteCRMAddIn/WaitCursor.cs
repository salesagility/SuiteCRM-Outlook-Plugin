using SuiteCRMClient.Logging;
using System;
using System.Windows.Forms;

namespace SuiteCRMAddIn
{
    public struct WaitCursor: IDisposable
    {
        private readonly Form _form;
        private readonly Cursor _originalCursor;
        private readonly bool _shouldReenable;

        public WaitCursor(Form form, bool shouldDisable = false)
        {
            _form = form;
            _originalCursor = form.Cursor;
            _shouldReenable = shouldDisable && form.Enabled;

            try
            {
                form.Cursor = Cursors.WaitCursor;
                if (_shouldReenable) form.Enabled = false;
            }
            catch (Exception any)
            {
                // doesn't, cosmically speaking, matter.
                Globals.ThisAddIn.Log.Warn($"Exception while trying to set wait cursor on form {_form.Name}", any);
            }
        }

        public void Dispose()
        {
            _form.Cursor = _originalCursor;
            if (_shouldReenable) _form.Enabled = true;
        }

        public static WaitCursor For(Form form, bool disableForm = false)
        {
            return new WaitCursor(form, disableForm);
        }
    }
}
