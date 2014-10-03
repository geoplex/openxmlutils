#region File Information
//
// File: "SpreadsheetUnitTest.cs"
// Purpose: "Some basic tests to demonstrate the Spreadsheet wrapper class"
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
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace OpenXmlUtils.Tests
{
    public class Song
    {
        public string Artist { get; set; }
        public string Title { get; set; }
        public double Double { get; set; }
        public long Int { get; set; }
        public bool Bool { get; set; }
        public DateTime Date { get; set; }
        public TimeSpan TimeSpan { get; set; }
    }

    [TestClass]
    public class SpreadsheetUnitTest
    {
        [TestMethod]
        public void TestObjectsToSpreadsheet()
        {
            var songs =
                    new List<Song>
                        { new Song { Artist = "Joy Devision", Title = "Disorder", Date = DateTime.Today, TimeSpan = TimeSpan.FromSeconds(3343), Int = 89453312L, Double = 4043.4545, Bool = false },
                          new Song { Artist = "Moderate", Title = "A New Error", Date = DateTime.Today, TimeSpan = TimeSpan.FromSeconds(34345), Int = 89563312L, Double = 5.6, Bool = true },
                          new Song { Artist = "Massive Attack", Title = "Paradise Circus", Date = DateTime.Today + TimeSpan.FromDays(53), TimeSpan = TimeSpan.FromSeconds(545), Int = 344334L, Double = 222.3, Bool = false },
                          new Song { Artist = "The Horrors", Title = "Still Life", Date = DateTime.Today - TimeSpan.FromDays(1), TimeSpan = TimeSpan.FromSeconds(22345), Int = 9497934L, Double = 33.4634444, Bool = true },
                          new Song { Artist = "Todd Terje", Title = "Inspector Norse", Date = DateTime.Today - TimeSpan.FromDays(356), TimeSpan = TimeSpan.FromSeconds(5565), Int = 34211343L, Double = 54.44444, Bool = false },
                          new Song { Artist = "Alpine", Title = "Hands", Date = DateTime.Today - TimeSpan.FromDays(5.5), TimeSpan = TimeSpan.FromSeconds(9907), Int = 32323333L, Double = 3445.44, Bool = false },
                          new Song { Artist = "Parquet Courts", Title = "Ducking and Dodging", Date = DateTime.Today - TimeSpan.FromDays(88.55), TimeSpan = TimeSpan.FromSeconds(8877), Int = 8088872L, Double = 44.0, Bool = false },        
                        };

            var fields = new List<SpreadsheetField>
            {
                new SpreadsheetField{ Title = "Artist", FieldName = "Artist"},
                new SpreadsheetField{ Title = "Title", FieldName = "Title"},
                new SpreadsheetField{ Title = "RandomDate", FieldName = "Date"},
                new SpreadsheetField{ Title = "RandomTimeSpan", FieldName = "TimeSpan"},
                new SpreadsheetField{ Title = "RandomInt", FieldName = "Int"},
                new SpreadsheetField{ Title = "RandomDouble", FieldName = "Double"},
                new SpreadsheetField{ Title = "RandomBool", FieldName = "Bool"}
            };

            Spreadsheet.Create(@"C:\temp\songs.xlsx",
                new SheetDefinition<Song>
                {
                    Fields = fields,
                    Name = "Songs",
                    SubTitle = DateTime.Today.ToLongDateString(),
                    IncludeTotalsRow = true,
                    Objects = songs
                });
        }

        [TestMethod]
        public void TestDictionariesToSpreadsheet()
        {
            var songs =
                new List<IDictionary<string, object>>
            {
                new Dictionary<string, object> { { "Artist" , "Joy Devision"}, {"Title" , "Disorder"}, {"Date" , DateTime.Today}, {"TimeSpan" , TimeSpan.FromSeconds(3343)}, {"Int" ,89453312L},{ "Double" , 4043.4545}, {"Bool" , false }},
                new Dictionary<string, object> { { "Artist" , "Moderate"}, {"Title" , "A New Error"}, {"Date" , DateTime.Today}, {"TimeSpan" , TimeSpan.FromSeconds(34345)}, {"Int", 89563312L},{ "Double" , 5.6}, {"Bool" , true }},
                new Dictionary<string, object> { { "Artist" , "Massive Attack"}, {"Title" , "Paradise Circus"}, {"Date" , DateTime.Today}, {"TimeSpan" , TimeSpan.FromSeconds(545)}, {"Int", 344334L},{ "Double" , 222.3}, {"Bool" , false }},
                new Dictionary<string, object> { { "Artist" , "The Horrors"}, {"Title" , "Still Life"}, {"Date" , DateTime.Today}, {"TimeSpan" , TimeSpan.FromSeconds(1123)}, {"Int", 9497934L},{ "Double" , 33.4634444}, {"Bool" , true }},
                new Dictionary<string, object> { { "Artist" , "Todd Terje"}, {"Title" , "Inspector Norse"}, {"Date" , DateTime.Today}, {"TimeSpan" , TimeSpan.FromSeconds(9973)}, {"Int", 34211343L},{ "Double" , 54.44444}, {"Bool" , false }},
                new Dictionary<string, object> { { "Artist" , "Alpine"}, {"Title" , "Hands"}, {"Date" , DateTime.Today}, {"TimeSpan" , TimeSpan.FromSeconds(3841)}, {"Int", 32323333L},{ "Double" , 3445.44}, {"Bool" , false }},
                new Dictionary<string, object> { { "Artist" , "Parquet Courts"}, {"Title" , "Ducking and Dodging"}, {"Date" , DateTime.Today}, {"TimeSpan" , TimeSpan.FromSeconds(9973)}, {"Int", 8088872L},{ "Double" , 44.0}, {"Bool" , false }}        
            };

            var fields = new List<SpreadsheetField>
            {
                new SpreadsheetField{ Title = "Artist", FieldName = "Artist"},
                new SpreadsheetField{ Title = "Title", FieldName = "Title"},
                new SpreadsheetField{ Title = "RandomDate", FieldName = "Date"},
                new SpreadsheetField{ Title = "RandomTimeSpan", FieldName = "TimeSpan"},
                new SpreadsheetField{ Title = "RandomInt", FieldName = "Int"},
                new SpreadsheetField{ Title = "RandomDouble", FieldName = "Double"},
                new SpreadsheetField{ Title = "RandomBool", FieldName = "Bool"}
            };

            Spreadsheet.Create(@"C:\temp\songs_dict.xlsx",
                new SheetDefinition<IDictionary<string, object>>
                {
                    Fields = fields,
                    Name = "Songs",
                    IncludeTotalsRow = true,
                    Objects = songs
                });
        }

        [TestMethod]
        public void TestMultipleSheets()
        {
            var songs =
                    new List<Song>
                        { new Song { Artist = "Joy Devision", Title = "Disorder", Date = DateTime.Today, TimeSpan = TimeSpan.FromSeconds(3434), Int = 89453312L, Double = 4043.4545, Bool = false },
                          new Song { Artist = "Moderate", Title = "A New Error", Date = DateTime.Today, TimeSpan = TimeSpan.FromSeconds(6576), Int = 89563312L, Double = 5.6, Bool = true },
                          new Song { Artist = "Massive Attack", Title = "Paradise Circus", Date = DateTime.Today + TimeSpan.FromDays(53), TimeSpan = TimeSpan.FromSeconds(9974), Int = 344334L, Double = 222.3, Bool = false },
                          new Song { Artist = "The Horrors", Title = "Still Life", Date = DateTime.Today - TimeSpan.FromDays(1), TimeSpan = TimeSpan.FromSeconds(9935), Int = 9497934L, Double = 33.4634444, Bool = true },
                        };

            var songs2 =
                    new List<Song>
                        { new Song { Artist = "Todd Terje", Title = "Inspector Norse", Date = DateTime.Today - TimeSpan.FromDays(356), TimeSpan = TimeSpan.FromSeconds(9009), Int = 34211343L, Double = 54.44444, Bool = false },
                          new Song { Artist = "Alpine", Title = "Hands", Date = DateTime.Today - TimeSpan.FromDays(5.5), TimeSpan = TimeSpan.FromSeconds(8836), Int = 32323333L, Double = 3445.44, Bool = false },
                          new Song { Artist = "Parquet Courts", Title = "Ducking and Dodging", Date = DateTime.Today - TimeSpan.FromDays(88.55), TimeSpan = TimeSpan.FromSeconds(1162), Int = 8088872L, Double = 44.0, Bool = false },        
                        };

            var fields = new List<SpreadsheetField>
            {
                new SpreadsheetField{ Title = "Artist", FieldName = "Artist"},
                new SpreadsheetField{ Title = "Title", FieldName = "Title"},
                new SpreadsheetField{ Title = "RandomDate", FieldName = "Date"},
                new SpreadsheetField{ Title = "RandomTimeSpan", FieldName = "TimeSpan"},
                new SpreadsheetField{ Title = "RandomInt", FieldName = "Int"},
                new SpreadsheetField{ Title = "RandomDouble", FieldName = "Double"},
                new SpreadsheetField{ Title = "RandomBool", FieldName = "Bool"}
            };

            Spreadsheet.Create(@"C:\temp\songs_multi.xlsx",
                new List<SheetDefinition<Song>>
                {
                    new SheetDefinition<Song>
                    {
                        Fields = fields,
                        Name = "1",
                        IncludeTotalsRow = true,
                        Objects = songs
                    },
                    new SheetDefinition<Song>
                    {
                        Fields = fields,
                        Name = "2",
                        IncludeTotalsRow = true,
                        Objects = songs2
                    }
                });
        }
    }
}
