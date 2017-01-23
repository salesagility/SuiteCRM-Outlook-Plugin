namespace SuiteCRMClient
{
    using System;

    public class clsAddValue
    {
        private string m_Display;
        private string m_Value;

        public clsAddValue(string Display, string Value)
        {
            this.m_Display = Display;
            this.m_Value = Value;
        }

        public override string ToString()
        {
            return this.Display;
        }

        public string Display
        {
            get
            {
                return this.m_Display;
            }
        }

        public string Value
        {
            get
            {
                return this.m_Value;
            }
        }
    }
}
