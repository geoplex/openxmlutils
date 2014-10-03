#region File Information
//
// File: "DateCell.cs"
// Purpose: "A simple class for date cells"
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

using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlUtils
{
    public class DateCell : Cell
    {
        public DateCell(string header, DateTime dateTime, int index)
        {
            DataType = CellValues.Date;
            CellReference = header + index;
            StyleIndex = (UInt32)CustomStylesheet.CustomCellFormats.DefaultDate;
            CellValue = new CellValue(dateTime.ToString("yyyy-MM-dd"));
        }
    }
}