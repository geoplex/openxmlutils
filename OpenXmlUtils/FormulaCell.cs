#region File Information
//
// File: "FormulaCell.cs"
// Purpose: "A simple class for formula cells"
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

using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlUtils
{
    public class FormulaCell : Cell
    {
        public FormulaCell(string header, string text, int index)
        {
            CellFormula = new CellFormula {CalculateCell = true, Text = text};
            DataType = CellValues.Number;
            CellReference = header + index;
        }
    }
}