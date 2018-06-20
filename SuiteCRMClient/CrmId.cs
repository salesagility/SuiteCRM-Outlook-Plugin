using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;

namespace SuiteCRMClient
{
    /// <summary>
    ///     A validated CRM id.
    /// </summary>
    public class CrmId : IComparable
    {
        public static readonly CrmId Empty = new CrmId();

        private static readonly Regex Validator = new Regex("[a-f0-9-]+");

        private static readonly Dictionary<string,CrmId> Issued = new Dictionary<string, CrmId>();

        /// <summary>
        ///     The actual id string.
        /// </summary>
        /// <remarks>
        ///     This class would specialise <see cref="string" /> only you can't.
        /// </remarks>
        private readonly string crmId;

        private CrmId()
        {
            crmId = string.Empty;
        }

        /// <summary>
        /// Create a new instance of a CrmId with this id.
        /// </summary>
        /// <remarks>
        /// This has to be public so that the JSON deserialiser can use it - but don't use it
        /// otherwise
        /// </remarks>
        /// <param name="id"></param>
        public CrmId(string id)
        {
            if (IsValid(id))
            {
                crmId = id;
                Issued[id] = this;
            }
            else
            {
                throw new TypeInitializationException($"'{id}' does not appear to be a valid CRM id.", null);
            }
        }

        public int CompareTo(object obj)
        {
            return crmId.CompareTo(obj.ToString());
        }

        public override string ToString()
        {
            return crmId;
        }

        /// <summary>
        ///     Validates a CRM id.
        /// </summary>
        /// <param name="id">The string which may or may not be a valid CRM id.</param>
        /// <returns>
        ///     True if `id` matches <see cref="Validator" /> pattern and
        ///     is of suitable length.
        /// </returns>
        public static bool IsValid(string id)
        {
            return !string.IsNullOrEmpty(id) && Validator.IsMatch(id) && id.Length == 36;
        }

        /// <summary>
        ///     Validates a CRM id.
        /// </summary>
        /// <param name="id">The object which may or may not be a valid CRM id.</param>
        /// <returns>
        ///     True if `id` is not null, matches <see cref="Validator" /> pattern and
        ///     is of suitable length.
        /// </returns>
        public static bool IsValid(CrmId id)
        {
            return id != null && id.IsValid();
        }

        /// <summary>
        ///     Validates a CRM id.
        /// </summary>
        /// <returns>
        ///     True if I match <see cref="Validator" /> pattern and
        ///     am of suitable length.
        /// </returns>
        public bool IsValid()
        {
            return IsValid(crmId);
        }

        /// <summary>
        /// True if <see cref="CrmId.IsValid"/> is false of this id.
        /// </summary>
        /// <param name="id">The object which may or may not be a valid CRM id.</param>
        /// <returns>True if <see cref="CrmId.IsValid"/> is false of this id.</returns>
        public static bool IsInvalid(CrmId id)
        {
            return !CrmId.IsValid(id);
        }

        public override bool Equals(object obj)
        {
            return base.Equals(obj) || (obj as CrmId)?.ToString() == crmId;
        }

        public override int GetHashCode()
        {
            return crmId.GetHashCode();
        }

        /// <summary>
        /// Get the single CrmId instance for this value.
        /// </summary>
        /// <param name="value">The value to seek.</param>
        /// <returns>A CrmId instance</returns>
        /// <exception cref="TypeInitializationException"> if `value` does not appear to be a valid CRM id.</exception>
        public static CrmId Get(string value)
        {
            return string.IsNullOrEmpty(value) ?
                CrmId.Empty :
                CrmId.Issued.ContainsKey(value) ?
                    CrmId.Issued[value]:
                    new CrmId(value);
        }

        /// <summary>
        /// Get the single CrmId instance for this value.
        /// </summary>
        /// <param name="value">The value to seek.</param>
        /// <returns>A CrmId instance</returns>
        /// <exception cref="TypeInitializationException"> if `value` does not appear to be a valid CRM id.</exception>
        public static CrmId Get(object value)
        {
            return value == null ? CrmId.Empty : CrmId.Get(value.ToString());
        }
    }
}