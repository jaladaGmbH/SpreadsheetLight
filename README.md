# SpreadsheetLight

Open source developer-friendly spreadsheet library compatible with Microsoft Excel 2007/2010/2013 and LibreOffice Calc

SpreadsheetLight is an open source Open XML spreadsheet library for .NET Framework written in C#, and is released under the MIT License. You can create new Open XML spreadsheets, or work with existing Open XML spreadsheets that are compatible with Microsoft Excel 2007/2010/2013 and LibreOffice Calc.

This is the fork from www.spreadsheetlight.com library with some improvements.

Library was merged with required old version of DocumentFormat.OpenXml to allow use new one.

####How to start
    
First add project to your solution.

Example:

    using (SLDocument sl = new SLDocument("Template.xlsx"))
    {
        if (sl.SelectWorksheet("Sheet1"))
        {
            sl.SetCellValue("A1", "Hello, world!");
        }
        sl.SaveAs("GeneratedReport.xlsx");
    }

You can find more tutorial and examples on [Tutorial](http://spreadsheetlight.com/tutorial/) and [Samples](http://spreadsheetlight.com/sample-code/) on [SpreadsheetLight site](www.spreadsheetlight.com).

## Changelog
**Version 3.5.0** (14 Nov 2020)
* Breaking change: Migrated from targetting .NET Framework 4 to .NET Standard 2
**Version 3.4.9** (1 Apr 2017)
* Get filter range of a worksheet if a filter exists. See SLDocument.GetFilterRange().
* On the Mono Framework, the error “System.OverflowException : Arithmetic operation resulted in an overflow.” no longer occurs. This was due to the System.Drawing.Bitmap not setting a default resolution (within Mono Framework. Running on .NET is fine. This part happens during the initialisation phase for setting the internal SLSimpleTheme class). Mono uses libgdiplus, which apparently assigns the Bitmap object zero for both horizontal and vertical resolutions. Hence the overflow error.
* Bug fix: It was not possible to copy a cell style from a cell to another cell that’s on the same row or column. It is now fixed (it was De Morgan’s laws incorrectly implemented).
* Bug fix: Creating a data validation that references another worksheet as a data source now works correctly.
* Bug fix: Calling InsertRow() multiple times now work.
* Bug fix: If a formula is passed as a parameter to HighlightCellsEqual() in SLConditionalFormatting class, it presents incorrect behaviour (something like $B$1 becomes “$B$1”, with incorrectly double quoting). It is now fixed (along HighlightCellsBeginningWith(), HighlightCellsBetween(), HighlightCellsContainingText(), HighlightCellsEndingWith(), HighlightCellsGreaterThan(), HighLightCellsLessThan())
* Bug fix: Copying a worksheet now returns true if successful.
* Bug fix: Deleting rows sometimes deleted incorrect rows. That was due to random rows being deleted (assigned by OS so it “looks” like it’s random). It is now fixed.
**Version 3.4.8** (22 Oct 2016)
* You can now check with SLDocument.HasCellError(). WARNING: SpreadsheetLight does not have a formula calculation engine, so only existing errors are reported.
* Bug fix: SLStyle.Alignment.TextRotation now rotates text correctly.
* Bug fix: Worksheet names are now (more? [Is that grammatically correct?]) correctly surrounded with single quotes in formulas and chart sheet formulas if they contain special characters (every special character on your keyboard except underscore and period).
* Bug fix: GetCellValueAsString() now no longer contain the phonetic text (if it exists). As part of the correction, SLRstType.ToPlainString() also no longer contain the phonetic text.
* Bug fix: Inline strings in existing Excel files will now be correctly loaded (and saved into the shared strings table if the particular worksheet with the inline strings is selected).
* Thanks to these awesome people for sending feature and bug requests: Bodo F, JF, HK, Vincent D, Jerry S

**Version 3.4.7** (10 Oct 2016)
* Bug fix: Selecting worksheets one after another that has row properties (such as custom row heights) will no longer cause an error. (The error is actually in WriteSelectedWorksheet(), where the iteration was over a combined list of row properties indices and cell row indices, and the error occurs if there’s a row properties row without any cells.)
Thanks to David L for telling me about the bug.

**Version 3.4.6** (3 Oct 2016)
* Shared string unique count is now written in the file. This makes opening the resulting Excel file faster if there are large number of text strings. Set the property SLDocument.WriteUniqueSharedStringCount to false if the file opens with an error.
* Bug fix: Copy rows/columns now does not cause a runtime error (there were 2 separate cell stores and the wrong one was used, hence the reference index not found).

**Version 3.4.5** (26 Sep 2016)
* SmartTags is now removed from consideration (not so smart now, are you? ;). Which means the code is now ready for Open XML SDK 2.5! And yes, it now works with Open XML SDK 2.5 (have I mentioned that? lol)
* Internal cell storage structure is revamped. It used to be a 1-dimensional dictionary with a 2-dimensional key, but is now a 2-dimensional dictionary with a 1-dimensional key. Tech explanation: Dictionary to Dictionary>. This is done because a 1-dimensional key is at most 2^31 hash entries, which cannot contain the theoretical 1048576 (2^20) rows and 16384 (2^14) columns per worksheet supported by latest versions of Excel.
* Bug fix: plotting separate data series on charts as different chart types now works (your typical combination chart such as a column chart with lines)
* Breaking change: SLDocument.GetCells() now returns a Dictionary>

**Version 3.4.4** (26 Apr 2014)
* Bug fix: Formulas will be correctly changed when columns are deleted (when the formula involves said deleted columns).
**Version 3.4.3** (1 Mar 2014)
* Bug fix: Cell value/formula set on a shared cell formula base cell now works correctly. For example, setting on C3 when C3 holds a shared cell formula will work.
* Bug fix: Setting filter on worksheet now works correctly. It failed to sort before (basically also need to set underlying defined name _xlnm._FilterDatabase. Gawdiggitty.)

**Version 3.4.2** (22 Feb 2014)
* You can now get row/column grouping levels. See GetRowGroupLevel() and GetColumnGroupLevel() functions.
* You can now get a list of the shared cell formulas in the currently selected worksheet. See GetSharedCellFormulas() function.
* Cell formulas are now more correctly maintained when copying/inserting/deleting cells/rows/columns.
* Catered for situation where cell reference ranges aren’t in top-left to bottom-right format (such as E1:A7). See SLTool for translating reference sequences to SLCellPointRange and vice versa. Don’t worry, you probably wouldn’t have known about this anyway…
* You can now draw borders on a cell range! And border grids! Try out the DrawBorder() and DrawBorderGrid() functions.
* You can now merge cells and set style/border properties at the same time! No more manual border drawing on merged cells! Woohoo! See MergeWorksheetCells() function overloads.
* You can now make SpreadsheetLight throw up, I mean, throw exceptions! When there are exceptions, of course. See ThrowExceptionsIfAny property. Not sure how much help this is though…
* You can now forcibly get a boolean value if it looks like it’s a boolean but actually stored as text. See the GetCellValueAsBoolean() functions.
* Breaking change: SLCellFormula.Reference data type changed from string to List (but you shouldn’t have been using this anyway…)
* Bug fix: The properties “count” and “uniqueCount” removed from shared strings table part. It seems a high number of shared strings will cause the calculation to render a corrupt file. See when writing the shared strings table for details.
* Bug fix: Outline (grouping) levels of rows and columns now limited to 0 to 7 (was allowed to go to 8 previously. See GroupRows() and GroupColumns() in RowColumnFunctions.cs)

**Version 3.4.1** (31 Aug 2013)
* Removed optional argument use (specifically IsStylish) so that developers using Visual Studio 2008 (.NET Framework 3.5) can still compile the source code.
* Optimised GetWorksheetStatistics(). It’s now faster and less memory intensive.
* Bug fix: autofitting rows with a smaller initial height now correctly autofit to content
* Bug fix (?): SLStyle.MergeStyle() now takes on the new style object’s format code regardless.
**Version 3.4** (27 July 2013)
* You can now import text! See ImportText() and its relevant SLTextImportOptions class for details.
* You can now make stock charts.
* 8 MOAR THEMES! Berlin, Circuit, Damask, Depth, Droplet, Main Event, Slate, Vapor Trail.
* Exposed the HasAutoFilter property for SLTable. Don’t want a filter for your table? Set it to false.
* Removed “using System;” for SLTuple.cs and SLTuplesType.cs. Because .NET Framework 4 has a System.Tuple, which clashes with DocumentFormat.OpenXml.Spreadsheet.Tuple.
* Built-in number format code index 14 now gets the format from computer’s regional short date settings (instead of just mm-dd-yyyy from Open XML specs).
* Charts now correctly set the auto-label when changing from category axis to date axis (or vice versa).
* Added NoMultiLevelLabels and ShowDataLabelsOverMaximum to charts.
* SLPicture now also uses EMF image files (basically with the Image class instead of the Bitmap class) to set the internal horizontal/vertical resolutions.
* Bug fix: Loading a spreadsheet with an existing chart sheet now work properly.
* Bug fix: Autofitting row/column now works when a large positive value is in a cell with default numeric format (it was mistaken as a date format).
* Bug fix: The totals row value now follow the corresponding row/column style (specifically the number format).
* Bug fix: The SLShapeProperties class will no longer render a default (ie. empty) SLReflection class.
* Special thanks to these marvelous people for suggestions and informing me of bugs:
* Thomas Z, David H, Chris K, Stefano L

**Version 3.3** (19 May 2013)
* You can now get comment text (but no comment box style properties. Sorry). See GetCommentText() of SLDocument class.
* You can now resize a picture using percentages of the original size (see SLPicture).
* Bug fix: Autofitting a column with forced new lines in the cell content now fits correctly.
* Bug fix: Autofitting without specifying font names now work correctly.
* Bug fix: Explicitly setting boolean font properties (SLFont) such as bold and italic to false now work correctly.
* Bug fix: Freezing only rows or only columns now work correctly.
* Bug fix: Inserting/deleting rows on a completely new worksheet after setting some cell values now work correctly.
* Bug fix: Appending text that starts with a number to the page header/footer now works correctly.
* Bug fix: PlotVisibleOnly is now always rendered.
* Bug fix: Setting a styled text in the header/footer where the text starts with a number is an error. The start of the text (which is a number) combines with the font size of the font setting and results in an unusually large font sized text (that’s not even correct). For example, Arial 12pt with “2/25/2013” becomes Arial 122 pt with “/25/2013”.
* Bug fix: Selecting a worksheet now returns true if the worksheet to be selected is already selected. (Ha! Unravel that convoluted statement!)
* Bug fix: Existing spreadsheets with a calculation chain (shared formulas) and subsequently have no more calculation formulas (maybe a cell formula got deleted) will now have the corresponding calculation chain part removed.
* Special thanks to these wonderful people for giving suggestions or pointing out bugs:
* György S, Brian M, Ed L, Nick W, Anele M, Philipp H, Joel S, Chris K, Larry.
* 
* Extra special thanks to drakex for hosting a version on NuGet because I was too daft to do it. Yes the NuGet versions are now updated by me. Bug fixed versions will now be more promptly delivered.

**Version 3.2** (6 Feb 2013)
* You can now have the collection of existing cells, styles and shared strings. In case you want to view data via LINQ.
* Added 21 new themes from Office 2013. We’re now up to 74 built-in themes. Woohoo!
* Inserting/deleting rows/columns now also correctly moves/resizes existing images.
* Will now use the actively selected worksheet already selected on an existing spreadsheet.
* Removing hyperlinks now remove cleanly (a cell range with a hyperlink is now completely cleanly removed).
* Bug fix: Calculation cells now loaded correctly (Clone the thing while loading dangit!)
**Version 3.1** (18 Jan 2013)
* Copying cells now also copy hyperlinks.
* Style setting behaviour now more in line with Excel style setting behaviour.
* Speed optimisations in style setting.

**Version 3.0** (13 Jan 2013)
* Excel 2010 specific conditional formatting! Specifically data bars and icon sets.
* Support for data validations.
* Show/hide worksheets. Did you know you can have worksheets “very hidden”?
* You can now set the active cell in the currently selected worksheet.
* You can now set the selected worksheet. No special functions, just use the SelectWorksheet() as normal. Last selected worksheet will be the uh, selected worksheet. Just like using Excel!
* Set the print area. See SetPrintArea() of SLDocument class.
* You can now show formulas. Nothing to do with the formula bar. This just means showing formulas instead of calculated results.
* You can now copy cells with just the value, formulas, style formatting.
* You can also now copy cells from another worksheet to your current worksheet. Make use of the paste options!
* You can now do fancy rich text manipulation. See the SLRstType class for more.
* You can now set the header/footer text in a very easy manner. See the SLPageSettings class for more, such as SetLeftHeaderText(“I’m left handed”).
* You can now group/ungroup rows/columns.
* You can now collapse/expand groups of rows/columns.
* You can now make hyperlinks on cells.
* You can now set the scope of defined names too.
* You can now get a list of existing defined names. Be warned: there are a lot of properties that might boggle your mind…
* SLTool/SLConvert .ToCellRange() and .ToCellReference() now correctly wraps worksheet names with spaces in single quotes.
* Added ShapeProperties to SLPicture. Now you can do more stuff with your images!
* Speed optimisations for autofitting.
* Splitting panes now even less tedious. There’s an overload that just asks if how many rows/columns and if you want the row/column headings to exist.
* Breaking change! The list of SLMergeCell classes returned is now of List instead of an array.
* Bug fix: SLEffectList now actually render the soft edges…
* Bug fix: VML drawings with embedded images now start with relationship IDs of “rId1” (mainly for cell comments). Excel chokes otherwise… *sigh*
* Bug fix: Copying worksheets now also has hyperlinks copied correctly.
* Bug fix: Switch rapidly between existing worksheets of existing spreadsheet and setting cell values in between, will now have the cell values saved. Forgot to use WorksheetPart.Worksheet.Save()… *sigh*
* Bug fix: Renaming worksheets now has chart data references updated too.
* Bug fix: Freeze panes now work correctly in Excel 2007 (without spewing errors).

**Version 2.5** (28 Nov 2012)
* There’s autofitting support for row heights and column widths. This is possibly the most sought after feature.
* You can now filter data.
* Removed static column name and reverse column name List and Dictionary lookup tables. This made the library infinitesimally slower but freed the library from being plagued by multi-thread or multi-processor or multi-opened-document or multi-whatever problems.
* Modified the hashing algorithm of SLCellPoint. Who knew this would affect the speed?

**Version 2.4.1** (13 Nov 2012)
* Bug fix: Data label positions now correctly displayed when assigned.

**Version 2.4** (13 Nov 2012)
* Sparklines! This is an Excel 2010 (and later?) specific feature.
* More themes! Now with a total of 53 built-in themes! These are from Microsoft Excel by the way…
* SLConvert class now has more converting functions! Try out the ToCellReference() and ToCellRange() functions.
* We have SetPatternFill()! Comments can also have a pattern fill!
* Charts can now glow and have soft edges!
* Added GetCurrentWorksheetName() function.
* Added GetSheetNames() function.
* We have a GetWorksheetStatistics() function with the corresponding SLWorksheetStatistics. Get the number of rows/columns/cells used in your worksheet! Don’t know what other stats are useful though…
* Various changes to make SLChart render as close to that of Excel 2010 as possible. This also fixes many conformance issues because apparently Excel 2007 is fairly tolerant but Excel 2010 has a big stick to whack any inconsistencies on the head. The specs say it’s optional; yeah optional my foot…
* Clone() function added to various classes. Notables are those of charting, conditional formatting, picture and styling.
* Breaking change: SpreadsheetLight.SLFill has the method signature of SetPattern() changed to more like that of SpreadsheetLight.Drawing.SLFill. Basically, the pattern type is in front of the foreground and background colours.
* Bug fix: Data labels don’t render correctly if category/series/value names are set to false.

**Version 2.3** (4 Nov 2012)
* Support for data point customisation in charts.
* Support for data labels in charts.
* Support for data tables in charts.
* Optimised writing of theme XML file. Shaved maybe 3 seconds off of total time. Is it worth it? You tell me…
* Overlay property of SLLegend class now works better.
* Code cleanup: “using” statements pared to … much fewer of them…
* URI relationship targets are now more in line with what Microsoft Excel creates. This probably has close to zero impact on you. Unless you use LibreOffice or iPhone/iPad.
* Bug fix: Error when loading in existing spreadsheet with calculation cells.

**Version 2.2** (23 Oct 2012)
* We have cell comments!
* Added overloads for SetCellValue() for SLRstType
* Chart axis labels can now be rotated and such.
* Any cell value set that starts with an equal sign “=” or a single quote “‘” is now better handled.
* You can now open an existing spreadsheet that came in a Stream object.
* It’s now easier to set a theme for the spreadsheet on initialisation.
* Add and delete background pictures.
* New class SLConvert for convenient miscellaneous convertions.
* Bug fix: When loading byte data for pictures, the byte data is now loaded by value. Previously, if you change the original byte data, the loaded picture data also change.
**Version 2.1.1** (9 Oct 2012)
Fixed URI thingie in relationship file so LibreOffice Calc can load in metadata. Ermahgerd fixing the URI thingie in a package is annoying… On the upside, we can still use .NET Framework 3.5 (instead of 4.0). For now…
**Version 2.1** (8 Oct 2012)
* Work with multiple spreadsheets at one go! (see sample source code page for examples)
* Full support for core document properties.
* Better styling support compatible with LibreOffice Calc.
* LibreOffice Calc can load spreadsheet files from SpreadsheetLight when document properties are set.
* Line data series now can be smoothed (line charts and scatter charts) independently of the selected chart type.
* ImportDataTable() now more properly implemented when DBNull data occurs. Much thanks to Troye Stonich.
* Added CloseWithoutSaving() function.
* Breaking change! Minimum .NET Framework version required is .NET Framework 4.0.
* Breaking change! Several classes now has to be created from SLDocument instead of using the “new” coding construct. The important ones are SLStyle, SLFont, SLTable and SLChart.

**Version 2.0** (29 Sep 2012)
* Speed and memory optimisation! Now with the speed and power of hurricanes yet with the memory footprint of a breeze. Achievement unlocked: Quad-Core Zephyr.
* Create your own combination charts! Plot data series on the primary axis! Or on the secondary axis! Warning: only very weak checks are done on whether your resulting combination chart is valid. But if Excel displays your combination chart just fine, then you’re fine too.
* Chart data series customisations. Colour individual columns differently if you’re so inclined…
* Axis title customisations. Rotate it, set bright purple font, italicise it, underline it. Go crazy! (actually don’t…)
* Bubble charts enabled! But I still don’t know how to use it correctly… Use with caution.
* Styling properties (fill, border, shadow, 3D format) for the chart area. But you know, be tasteful…
* Support for customising the floor, side wall and back wall of 3D charts (satisfy your inner home making instincts)
* Worksheet protection (but no password protection)
* Insert and remove page breaks
* Change the page layout (Normal, Page Layout, Page Break Preview)
* Added ClearCellContent() overload with no parameters, clearing all cell content in the currently selected worksheet.
* Added the handling of System.DBNull for importing data tables.
* Breaking change: You can’t (or shouldn’t) set ShowHiddenData property of SLChart as and when you like. It’s now done together when creating a new instance of SLChart. Because we need to know if hidden data is included at the start of creating the chart. This has a 0.001% chance of affecting you…
* Bug fix: to always remove existing relevant child elements for existing worksheet (in case there were existing say merged cells, but we unmerged all cells, then there’d be a lingering XML child element, but we don’t have any custom library classes for merged cells. Too long to explain…)
* Bug fix: CellStylesFormat forced save failed when there are duplicates (when opening existing file written with such duplicates)
* Bug fix: SLColorTransform class is now properly cloned. This would’ve caused themed colours to not work properly outside of stylesheets. Also, the tint wasn’t properly implemented.
* Bug fix: DateTime’s now properly calculated for GetCellValueAsDateTime() functions.
* Bug fix: SetRowHeight() failed when your regional settings use the comma instead of the decimal point. CultureInfo.InvariantCulture! I have dug through all ToString()’s of integral and floating point variables and added cultural insensitivity…
* Bug fix: Border styles persisted between SLStyle class variables, due to incorrect SLBorder initialisation.
* Version 1.2 (25 Aug 2012)
* SLPicture can now load image data in byte array form (thanks to Rob Hutchinson for his code submission and suggestion!)
* Bulk upload of data using DataTable
* Automagic XML escaping for string data (I’m looking at your Mr Ampersand…)
* Cultural indifference for numeric data. But the style format code still has to be in invariant culture mode.
* Fixed assignment to SLLine3DChart of SLPlotArea when SLLineChartType.Line3D used (used SLLineChart before). This didn’t create a bug but the wrong assignment bugs the heck out of me…
* Modified HideChartTitle() of SLChart to use AutoTitleDeleted. This will handle the auto-title for pie charts.

**Version 1.1.7.1** (17 Jul 2012)
Fixed bug. Will crash when gradient fill of SLShapeProperties is used (SLGradientStop not properly initialised when given a hexadecimal colour value)
Added functions to group several settings together for: top bevels, bottom bevels, extrusion and contour on the SLFormat3D class
Added a function to allow hiding of the chart legend

**Version 1.1.7** (16 Jul 2012)
* Breaking Change! New namespaces introduced: Charts and Drawing.

* SLChart class is now in Charts namespace (as with all the charting related classes)
* SLPicture class is now in Drawing namespace
* SLPictureJoinType enumeration is renamed SLLineJoinValues enumeration
* Insert charts as a chartsheet
* Styling for chart title, chart legend, and chart plot area
* Printing and page settings (page margins, header/footer, sheet tab colour and so on)

**Version 1.1.6** (2 Jun 2012)
* Basic chart support (no bubble nor stock charts. Bubble charts are weird…)
* Conditional formatting with formulas
* Relative position for pictures without forcing worksheet row/column dimensions

**Version 1.1.5.1** (24 May 2012)
* Fixed bug when renaming worksheets with an existing sheet name

**Version 1.1.5** (23 May 2012)
* Added column name function overloads (in addition to column index functions)
* Added IDisposable interface to SLDocument (so you can use “using” [or “Using” if VB.NET])
* Added conditional formatting for cells
* Added sorting capabilities (for tables and worksheet)
* Added table support (complete with subtotal functions)

**Version 1.1.4** (14 May 2012)
* Fixed bug on SLColor not having colours appearing (SetAllNull())
* Added copying, deleting and moving worksheets
* Added copying of styles from rows, columns and cells
* Added copying of rows, columns and cells
* More overloads of shortcut functions from font, fill, alignment onto SLStyle

**Version 1.1.3** (11 May 2012)
*Fixed bug on setting cells and overwriting any existing cell
*Set it such that only when image insertion with relative position or splitting forces custom row/column dimensions (this makes the worksheet look more “natural”)
*SetCellValueByRef() overloaded into SetCellValue() (34 overloaded functions!)
*Minor formula change on setting string cell values (check if equal sign at start, then set as numeric formula)

**Version 1.1.2** (1 May 2012)
* Used a 2-int structure instead of string for cell reference (sped up performance and reduced memory use. And made it easier internally to write code…)
* OpenXmlReader and OpenXmlWriter everywhere where appropriate (run faster!!)

**Version 1.1.1** (27 Apr 2012)
* Switched from ArrayList to List<>, Hashtable to Dictionary<>
* Removed custom list classes and used List<> instead
* Added strong name signing

**Version 1.1** (19 Feb 2012)
* Split panes!
* Insert/delete rows/columns (correctly accounting for cells, formulas, merged cells, tables, defined names, but not pictures/worksheet drawings. Drawings operate on a different “dimension”…)
* Simple defined named functionality
* Clear cell/row/column data
* Used a buffer workbook/worksheet for writing data, so the sheet?.xml files don’t always jump a number every time a new worksheet is added.
* Bug fix on unfreezing panes (different workbook views caused problems)
* Bug fix on target resolutions of SLPicture

**Version 1.0** (10 Jan 2012)
* Basic cell manipulation
* Basic worksheet manipulation
* Basic styling
* Shared string support
* Rich text support
* Basic theme support
* Merge cell support
* Picture support (we have 3D!)

## License

The MIT License (MIT)

Copyright (c) 2011 Vincent Tan Wai Lip

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
