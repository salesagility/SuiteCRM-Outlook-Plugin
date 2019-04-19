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
namespace SuiteCRMAddIn.BusinessLogic
{
    /// <summary>
    /// A direction in which things may be synchronised. Directions are named
    /// with respect to CRM.
    /// </summary>
    public class SyncDirection
    {
        /// <summary>
        /// The actual directions
        /// </summary>
        public enum Direction
        {
            Neither = 0,
            Export = 1,
            Import = 2,
            BiDirectional = 3
        }

        /// <summary>
        /// Convert this direction into a human-readable string
        /// </summary>
        /// <param name="direction">The direction.</param>
        /// <returns>The string.</returns>
        public static string ToString(Direction direction)
        {
            string result;

            switch (direction)
            {
                case Direction.Neither:
                    result = "None";
                    break;
                case Direction.Export:
                    result = "From CRM to Outlook";
                    break;
                case Direction.Import:
                    result = "From Outlook to CRM";
                    break;
                case Direction.BiDirectional:
                    result = "Both";
                    break;
                default:
                    result = "Shouldn't happen";
                    break;
            }

            return result;
        }

        /// <summary>
        /// Does this direction allow inbound transfers to Outlook?
        /// </summary>
        /// <param name="direction">The direction.</param>
        /// <returns>True if this direction allows inbound transfers.</returns>
        public static bool AllowInbound(Direction direction)
        {
            return (direction == Direction.Export || direction == Direction.BiDirectional);
        }

        /// <summary>
        /// Does this direction allow outbound transfers from Outlook?
        /// </summary>
        /// <param name="direction">The direction.</param>
        /// <returns>True if this direction allows outbound transfers.</returns>
        public static bool AllowOutbound(Direction direction)
        {
            return (direction == Direction.Import || direction == Direction.BiDirectional);
        }
    }
}
