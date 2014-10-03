
#region File Information
//
// File: "SheetDefinition.cs"
// Purpose: "Defines a single sheet (or tab) in a xlxs spreadsheet."
// Author: "Geoplex"
// 
#endregion

#region (c) Copyright 2014 Geoplex
//
// THE SOFTWARE IS PROVIDED "AS-IS" AND WITHOUT WARRANTY OF ANY KIND,
// EXPRESS, IMPLIED OR OTHERWISE, INCLUDING WITHOUT LIMITATION, ANY
// WARRANTY OF MERCHANTABILITY OR FITNESS FOR A PARTICULAR PURPOSE.
//
// IN NO EVENT SHALL GEOPLEX BE LIABLE FOR ANY SPECIAL, INCIDENTAL,
// INDIRECT OR CONSEQUENTIAL DAMAGES OF ANY KIND, OR ANY DAMAGES WHATSOEVER
// RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER OR NOT ADVISED OF THE
// POSSIBILITY OF DAMAGE, AND ON ANY THEORY OF LIABILITY, ARISING OUT OF OR IN
// CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.
//
#endregion

using System.Collections.Generic;

namespace OpenXmlUtils
{
    public class SheetDefinition<T>
    {
        /// <summary>
        /// Name of the sheet (shown in the tab)
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Title of the sheet
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Subtitle of the sheet
        /// </summary>
        public string SubTitle { get; set; }

        /// <summary>
        /// Objects to display in the sheet
        /// </summary>
        public IList<T> Objects { get; set; }

        /// <summary>
        /// Field names to extract from the objects and use as header names
        /// </summary>
        public List<SpreadsheetField> Fields { get; set; }

        /// <summary>
        /// Whether or not to include a row of calculated totals to the table
        /// </summary>
        public bool IncludeTotalsRow { get; set; }

    }
}
