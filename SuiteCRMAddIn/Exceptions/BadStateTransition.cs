/**
* Outlook integration for SuiteCRM.
* @package Outlook integration for SuiteCRM
* @copyright SalesAgility Ltd http://www.salesagility.com
*
* This program is free software; you can redistribute it and/or modify
* it under the terms of the GNU LESSER GENERAL PUBLIC LICENCE as published by
* the Free Software Foundation; either version 3 of the License, or
* (at your option) any later version.
*
* This program is distributed in the hope that it will be useful,
* but WITHOUT ANY WARRANTY; without even the implied warranty of
* MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
* GNU General Public License for more details.
*
* You should have received a copy of the GNU LESSER GENERAL PUBLIC LICENCE
* along with this program; if not, see http://www.gnu.org/licenses
* or write to the Free Software Foundation,Inc., 51 Franklin Street,
* Fifth Floor, Boston, MA 02110-1301  USA
*
* @author SalesAgility <info@salesagility.com>
*/
namespace SuiteCRMAddIn.Exceptions
{
    using BusinessLogic;
    using System;
    using System.Runtime.Serialization;

    /// <summary>
    /// An exception thrown when an unexpected transition is made on a syncstate.
    /// </summary>
    [Serializable]
    internal class BadStateTransition : NeverShowUserException
    {
        public TransmissionState From { get; private set; } = TransmissionState.Invalid;

        public TransmissionState To { get; private set; } = TransmissionState.Invalid;

        public BadStateTransition()
        {
        }

        public BadStateTransition(SyncState state, TransmissionState from, TransmissionState to) : base($"{state.GetType().Name}: failed transition {from} => {to}") {
            this.To = to;
            this.From = from;
        }
    }
}