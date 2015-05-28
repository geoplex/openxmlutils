OpenXMLUtils
============

Utility classes for the [Open XML SDK 2.5 for Office](http://msdn.microsoft.com/en-us/library/office/bb448854.aspx).

Copyright (c) Geoplex  All rights reserved.
Licensed under the Apache License, Version 2.0.
See License.txt in the project root for license information.

Package
=======

https://www.nuget.org/packages/OpenXmlUtils/

Documentation
=============

Currently supports creation of simple spreadsheets from a collection of objects based on this:
http://www.codeproject.com/Articles/97307/Using-C-and-Open-XML-SDK-for-Microsoft-Office. 

##Example Usage

Using a list of objects:
```c#
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
```

Using a list of dictionaries:
```c#
var songs =
	new List<object>
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
	new SheetDefinition<object>
	{
		Fields = fields,
		Name = "Songs",
		IncludeTotalsRow = true,
		Objects = songs
	});
```

Creating multiple sheets:
```c#
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
```

Hyperlinks:
```
var songs =
    new List<Song>
        { 
            new Song { Artist = "Parquet Courts", Title = "Ducking and Dodging", Url = "https://parquetcourts.wordpress.com", Hyperlink = "parquetcourts.wordpress.com"}, 
        };

var fields = new List<SpreadsheetField>
{
    new SpreadsheetField{ Title = "Artist", FieldName = "Artist"},
    new SpreadsheetField{ Title = "Title", FieldName = "Title"},
    new HyperlinkField{ Title = "Website", FieldName = "Url", DisplayFieldName = "Hyperlink"}
};

Spreadsheet.Create(@"C:\temp\songs_hyperlinks.xlsx",
    new SheetDefinition<Song>
    {
        Fields = fields,
        Name = "Songs",
        SubTitle = DateTime.Today.ToLongDateString(),
        IncludeTotalsRow = true,
        Objects = songs
    });
```

Row Grouping:
```c#
var songs =
    new List<object>
{
    new Dictionary<string, object> { { "Artist" , "Joy Devision"} },
    new List<object> {
            new Dictionary<string, object> { { "Albumn" , "Closer"} },
            new List<object> {
                new Dictionary<string, object>{ {"Title" , "Isolation"} },
                new Dictionary<string, object>{ {"Title" , "Colony"} },
                new Dictionary<string, object>{ {"Title" , "Decades"} },
            },
            new Dictionary<string, object> { { "Albumn" , "Unknown Pleasures"} },
            new List<object> {
                new Dictionary<string, object>{ {"Title" , "Disorder"} },
                new Dictionary<string, object>{ {"Title" , "Candidate"} },
                new Dictionary<string, object>{ {"Title" , "She's Lost Control"} }
            },
        },
    new Dictionary<string, object> { { "Artist" , "Moderate"} },
    new List<object> {
            new Dictionary<string, object>{ {"Title" , "A New Error"} },
            new Dictionary<string, object>{ {"Title" , "Rusty Nails"} },
            new Dictionary<string, object>{ {"Title" , "Seamonkey"} },
        },
};

var fields = new List<SpreadsheetField>
{
    new SpreadsheetField{ Title = "Artist", FieldName = "Artist"},
    new SpreadsheetField{ Title = "Albumn", FieldName = "Albumn"},
    new SpreadsheetField{ Title = "Title", FieldName = "Title"},
};

Spreadsheet.Create(@"C:\temp\songs_row_grouping.xlsx",
    new SheetDefinition<object>
    {
        Fields = fields,
        Name = "Songs",
        IncludeTotalsRow = false,
        Objects = songs
    });
```
