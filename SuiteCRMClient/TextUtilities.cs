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

namespace SuiteCRMClient
{
    using System.Text.RegularExpressions;

    /// <summary>
    /// Utility methods for mangling text.
    /// </summary>
    public class TextUtilities
    {
        /// <summary>
        /// If the string to modify contains the string to seek (ignoring differences in line ends)
        /// return that part of the string to modify which precedes the string to seek; otherwise
        /// return the string to modify unmodified. THIS IS NASTY. 
        /// </summary>
        /// <remarks>
        /// TODO: Move to a utility class.
        /// </remarks>
        /// <param name="toModify">The string which may be modified.</param>
        /// <param name="toSeek">The string to seek.</param>
        /// <returns>that part of the string to modify which precedes the string to seek.</returns>
        public static string StripAndTruncate(string toModify, string toSeek)
        {
            string result;

            if (string.IsNullOrEmpty(toSeek))
            {
                result = toModify;
            }
            else if (string.IsNullOrWhiteSpace(toModify))
            {
                result = string.Empty;
            }
            else
            {
                var offset = IndexIgnoreLineEnds(toModify, toSeek);
                var prefix = offset == -1 ?
                    toModify :
                    StripReturns(toModify).Substring(0, offset);

                result = Regex.Replace(prefix, @"\s+$", string.Empty);
            }

            return result;
        }

        /// <summary>
        /// Find the index of the string to seek in the string to search, ignoring differences in line ends.
        /// </summary>
        /// <remarks>
        /// This is obviously impossible since differences in line ends result in differences in offset; this 
        /// method treats any concatenation of potential line-end characters as a single character.
        /// </remarks>
        /// <param name="toSearch">The string to be searched.</param>
        /// <param name="toSeek">The string to seek.</param>
        /// <returns>An approximation of the index, or -1 if not found.</returns>
        private static int IndexIgnoreLineEnds(string toSearch, string toSeek)
        {
            string strippedSearch = Regex.Replace(string.IsNullOrEmpty(toSearch) ? string.Empty : toSearch, @" *[\n\r]+", ".");
            string strippedSeek = Regex.Replace(string.IsNullOrEmpty(toSeek) ? string.Empty : toSeek, @" *[\n\r]+", ".");

            return strippedSearch.IndexOf(strippedSeek);
        }

        /// <summary>
        /// Remove carriage return characters from this string.
        /// </summary>
        /// <param name="input">The string, which may contain carriage return characters.</param>
        /// <returns>A similar string, which does not. If input is null, return an empty string.</returns>
        private static string StripReturns(string input)
        {
            return string.IsNullOrWhiteSpace(input) ?
                string.Empty :
                Regex.Replace(input, @" *\r", "");
        }
    }
}
