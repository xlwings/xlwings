/**
 * The MIT License (MIT)
 * Copyright (c) Microsoft Corporation
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this software and 
 * associated documentation files (the "Software"), to deal in the Software without restriction, 
 * including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, 
 * and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, 
 * subject to the following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in all copies or substantial 
 * portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT 
 * NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. 
 * IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, 
 * WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE 
 * SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 * 
 * Source: https://github.com/OfficeDev/office-scripts-docs-reference/blob/main/generate-docs/script-inputs/excel.d.ts
*/

declare namespace ExcelScript {
    /*
     * Special Run Function
     */
    function run(
        callback: (workbook: Workbook) => Promise<void>
    ): Promise<void>;

    //
    // Class
    //

    /**
     * Contains information about a linked workbook.
     * If a workbook has links pointing to data in another workbook, the second workbook is linked to the first workbook.
     * In this scenario, the second workbook is called the "linked workbook".
     */
    interface LinkedWorkbook {
        /**
         * Makes a request to break the links pointing to the linked workbook.
         * Links in formulas are replaced with the latest fetched data.
         * The current `LinkedWorkbook` object is invalidated and removed from `LinkedWorkbookCollection`.
         */
        breakLinks(): void;

        /**
         * Makes a request to refresh the data retrieved from the linked workbook.
         */
        refreshLinks(): void;
    }

    /**
     * Represents the Excel application that manages the workbook.
     */
    interface Application {
        /**
         * Returns the Excel calculation engine version used for the last full recalculation.
         */
        getCalculationEngineVersion(): number;

        /**
         * Returns the calculation mode used in the workbook, as defined by the constants in `ExcelScript.CalculationMode`. Possible values are: `Automatic`, where Excel controls recalculation; `AutomaticExceptTables`, where Excel controls recalculation but ignores changes in tables; `Manual`, where calculation is done when the user requests it.
         */
        getCalculationMode(): CalculationMode;

        /**
         * Returns the calculation mode used in the workbook, as defined by the constants in `ExcelScript.CalculationMode`. Possible values are: `Automatic`, where Excel controls recalculation; `AutomaticExceptTables`, where Excel controls recalculation but ignores changes in tables; `Manual`, where calculation is done when the user requests it.
         */
        setCalculationMode(calculationMode: CalculationMode): void;

        /**
         * Returns the calculation state of the application. See `ExcelScript.CalculationState` for details.
         */
        getCalculationState(): CalculationState;

        /**
         * Provides information based on current system culture settings. This includes the culture names, number formatting, and other culturally dependent settings.
         */
        getCultureInfo(): CultureInfo;

        /**
         * Gets the string used as the decimal separator for numeric values. This is based on the local Excel settings.
         */
        getDecimalSeparator(): string;

        /**
         * Returns the iterative calculation settings.
         * In Excel on Windows and Mac, the settings will apply to the Excel Application.
         * In Excel on the web and other platforms, the settings will apply to the active workbook.
         */
        getIterativeCalculation(): IterativeCalculation;

        /**
         * Gets the string used to separate groups of digits to the left of the decimal for numeric values. This is based on the local Excel settings.
         */
        getThousandsSeparator(): string;

        /**
         * Specifies if the system separators of Excel are enabled.
         * System separators include the decimal separator and thousands separator.
         */
        getUseSystemSeparators(): boolean;

        /**
         * Recalculate all currently opened workbooks in Excel.
         * @param calculationType Specifies the calculation type to use. See `ExcelScript.CalculationType` for details.
         */
        calculate(calculationType: CalculationType): void;
    }

    /**
     * Represents the iterative calculation settings.
     */
    interface IterativeCalculation {
        /**
         * True if Excel will use iteration to resolve circular references.
         */
        getEnabled(): boolean;

        /**
         * True if Excel will use iteration to resolve circular references.
         */
        setEnabled(enabled: boolean): void;

        /**
         * Specifies the maximum amount of change between each iteration as Excel resolves circular references.
         */
        getMaxChange(): number;

        /**
         * Specifies the maximum amount of change between each iteration as Excel resolves circular references.
         */
        setMaxChange(maxChange: number): void;

        /**
         * Specifies the maximum number of iterations that Excel can use to resolve a circular reference.
         */
        getMaxIteration(): number;

        /**
         * Specifies the maximum number of iterations that Excel can use to resolve a circular reference.
         */
        setMaxIteration(maxIteration: number): void;
    }

    /**
     * Workbook is the top level object which contains related workbook objects such as worksheets, tables, and ranges.
     */
    interface Workbook {
        /**
         * Represents the Excel application instance that contains this workbook.
         */
        getApplication(): Application;

        /**
         * Specifies if the workbook is in AutoSave mode.
         */
        getAutoSave(): boolean;

        /**
         * Returns a number about the version of Excel Calculation Engine.
         */
        getCalculationEngineVersion(): number;

        /**
         * True if all charts in the workbook are tracking the actual data points to which they are attached.
         * False if the charts track the index of the data points.
         */
        getChartDataPointTrack(): boolean;

        /**
         * True if all charts in the workbook are tracking the actual data points to which they are attached.
         * False if the charts track the index of the data points.
         */
        setChartDataPointTrack(chartDataPointTrack: boolean): void;

        /**
         * Specifies if changes have been made since the workbook was last saved.
         * You can set this property to `true` if you want to close a modified workbook without either saving it or being prompted to save it.
         */
        getIsDirty(): boolean;

        /**
         * Specifies if changes have been made since the workbook was last saved.
         * You can set this property to `true` if you want to close a modified workbook without either saving it or being prompted to save it.
         */
        setIsDirty(isDirty: boolean): void;

        /**
         * Gets the workbook name.
         */
        getName(): string;

        /**
         * Specifies if the workbook has ever been saved locally or online.
         */
        getPreviouslySaved(): boolean;

        /**
         * Gets the workbook properties.
         */
        getProperties(): DocumentProperties;

        /**
         * Returns the protection object for a workbook.
         */
        getProtection(): WorkbookProtection;

        /**
         * Returns `true` if the workbook is open in read-only mode.
         */
        getReadOnly(): boolean;

        /**
         * True if calculations in this workbook will be done using only the precision of the numbers as they're displayed.
         * Data will permanently lose accuracy when switching this property from `false` to `true`.
         */
        getUsePrecisionAsDisplayed(): boolean;

        /**
         * True if calculations in this workbook will be done using only the precision of the numbers as they're displayed.
         * Data will permanently lose accuracy when switching this property from `false` to `true`.
         */
        setUsePrecisionAsDisplayed(usePrecisionAsDisplayed: boolean): void;

        /**
         * Gets the currently active cell from the workbook.
         */
        getActiveCell(): Range;

        /**
         * Gets the currently active chart in the workbook. If there is no active chart, then this method returns `undefined`.
         */
        getActiveChart(): Chart;

        /**
         * Gets the currently active slicer in the workbook. If there is no active slicer, then this method returns `undefined`.
         */
        getActiveSlicer(): Slicer;

        /**
         * Gets the currently selected single range from the workbook. If there are multiple ranges selected, this method will throw an error.
         */
        getSelectedRange(): Range;

        /**
         * Gets the currently selected one or more ranges from the workbook. Unlike `getSelectedRange()`, this method returns a `RangeAreas` object that represents all the selected ranges.
         */
        getSelectedRanges(): RangeAreas;

        /**
         * Represents a collection of bindings that are part of the workbook.
         */
        getBindings(): Binding[];

        /**
         * Add a new binding to a particular Range.
         * @param range Range to bind the binding to. May be a `Range` object or a string. If string, must contain the full address, including the sheet name
         * @param bindingType Type of binding. See `ExcelScript.BindingType`.
         * @param id Name of the binding.
         */
        addBinding(
            range: Range | string,
            bindingType: BindingType,
            id: string
        ): Binding;

        /**
         * Add a new binding based on a named item in the workbook.
         * If the named item references to multiple areas, the `InvalidReference` error will be returned.
         * @param name Name from which to create binding.
         * @param bindingType Type of binding. See `ExcelScript.BindingType`.
         * @param id Name of the binding.
         */
        addBindingFromNamedItem(
            name: string,
            bindingType: BindingType,
            id: string
        ): Binding;

        /**
         * Add a new binding based on the current selection.
         * If the selection has multiple areas, the `InvalidReference` error will be returned.
         * @param bindingType Type of binding. See `ExcelScript.BindingType`.
         * @param id Name of the binding.
         */
        addBindingFromSelection(bindingType: BindingType, id: string): Binding;

        /**
         * Gets a binding object by ID. If the binding object does not exist, then this method returns `undefined`.
         * @param id ID of the binding object to be retrieved.
         */
        getBinding(id: string): Binding | undefined;

        /**
         * Represents a collection of comments associated with the workbook.
         */
        getComments(): Comment[];

        /**
         * Creates a new comment with the given content on the given cell. An `InvalidArgument` error is thrown if the provided range is larger than one cell.
         * @param cellAddress The cell to which the comment is added. This can be a `Range` object or a string. If it's a string, it must contain the full address, including the sheet name. An `InvalidArgument` error is thrown if the provided range is larger than one cell.
         * @param content The comment's content. This can be either a string or `CommentRichContent` object. Strings are used for plain text. `CommentRichContent` objects allow for other comment features, such as mentions.
         * @param contentType Optional. The type of content contained within the comment. The default value is enum `ContentType.Plain`.
         */
        addComment(
            cellAddress: Range | string,
            content: CommentRichContent | string,
            contentType?: ContentType
        ): Comment;

        /**
         * Gets a comment from the collection based on its ID.
         * @param commentId The identifier for the comment.
         */
        getComment(commentId: string): Comment;

        /**
         * Gets the comment from the specified cell.
         * @param cellAddress The cell which the comment is on. This can be a `Range` object or a string. If it's a string, it must contain the full address, including the sheet name. An `InvalidArgument` error is thrown if the provided range is larger than one cell.
         */
        getCommentByCell(cellAddress: Range | string): Comment;

        /**
         * Gets the comment to which the given reply is connected.
         * @param replyId The identifier of comment reply.
         */
        getCommentByReplyId(replyId: string): Comment;

        /**
         * Represents the collection of custom XML parts contained by this workbook.
         */
        getCustomXmlParts(): CustomXmlPart[];

        /**
         * Adds a new custom XML part to the workbook.
         * @param xml XML content. Must be a valid XML fragment.
         */
        addCustomXmlPart(xml: string): CustomXmlPart;

        /**
         * Gets a new collection of custom XML parts whose namespaces match the given namespace.
         * @param namespaceUri This must be a fully qualified schema URI; for example, "http://schemas.contoso.com/review/1.0".
         */
        getCustomXmlPartsByNamespace(namespaceUri: string): CustomXmlPart[];

        /**
         * Gets a custom XML part based on its ID.
         * If the `CustomXmlPart` does not exist, then this method returns `undefined`.
         * @param id ID of the object to be retrieved.
         */
        getCustomXmlPart(id: string): CustomXmlPart | undefined;

        /**
         * Returns a collection of linked workbooks. In formulas, the workbook links can be used to reference data (cell values and names) outside of the current workbook.
         */
        getLinkedWorkbooks(): LinkedWorkbook[];

        /**
         * Represents the update mode of the workbook links. The mode is same for all of the workbook links present in the workbook.
         */
        getLinkedWorkbookRefreshMode(): WorkbookLinksRefreshMode;

        /**
         * Represents the update mode of the workbook links. The mode is same for all of the workbook links present in the workbook.
         */
        setLinkedWorkbookRefreshMode(
            linkedWorkbookRefreshMode: WorkbookLinksRefreshMode
        ): void;

        /**
         * Breaks all the links to the linked workbooks. Once the links are broken, any formulas referencing workbook links are removed entirely and replaced with the most recently retrieved values.
         */
        breakAllLinksToLinkedWorkbooks(): void;

        /**
         * Gets information about a linked workbook by its URL. If the workbook does not exist, then this method returns `undefined`.
         * @param key The URL of the linked workbook.
         */
        getLinkedWorkbookByUrl(key: string): LinkedWorkbook | undefined;

        /**
         * Makes a request to refresh all the workbook links.
         */
        refreshAllLinksToLinkedWorkbooks(): void;

        /**
         * Represents a collection of workbook-scoped named items (named ranges and constants).
         */
        getNames(): NamedItem[];

        /**
         * Adds a new name to the collection of the given scope.
         * @param name The name of the named item.
         * @param reference The formula or the range that the name will refer to.
         * @param comment Optional. The comment associated with the named item.
         */
        addNamedItem(
            name: string,
            reference: Range | string,
            comment?: string
        ): NamedItem;

        /**
         * Adds a new name to the collection of the given scope using the user's locale for the formula.
         * @param name The name of the named item.
         * @param formula The formula in the user's locale that the name will refer to.
         * @param comment Optional. The comment associated with the named item.
         */
        addNamedItemFormulaLocal(
            name: string,
            formula: string,
            comment?: string
        ): NamedItem;

        /**
         * Gets a `NamedItem` object using its name. If the object does not exist, then this method returns `undefined`.
         * @param name Nameditem name.
         */
        getNamedItem(name: string): NamedItem | undefined;

        /**
         * Represents a collection of PivotTableStyles associated with the workbook.
         */
        getPivotTableStyles(): PivotTableStyle[];

        /**
         * Creates a blank `PivotTableStyle` with the specified name.
         * @param name The unique name for the new PivotTable style. Will throw an `InvalidArgument` error if the name is already in use.
         * @param makeUniqueName Optional. Defaults to `false`. If `true`, will append numbers to the name in order to make it unique, if needed.
         */
        addPivotTableStyle(
            name: string,
            makeUniqueName?: boolean
        ): PivotTableStyle;

        /**
         * Gets the default PivotTable style for the parent object's scope.
         */
        getDefaultPivotTableStyle(): PivotTableStyle;

        /**
         * Gets a `PivotTableStyle` by name. If the `PivotTableStyle` does not exist, then this method returns `undefined`.
         * @param name Name of the PivotTable style to be retrieved.
         */
        getPivotTableStyle(name: string): PivotTableStyle | undefined;

        /**
         * Sets the default PivotTable style for use in the parent object's scope.
         * @param newDefaultStyle The `PivotTableStyle` object, or name of the `PivotTableStyle` object, that should be the new default.
         */
        setDefaultPivotTableStyle(
            newDefaultStyle: PivotTableStyle | string
        ): void;

        /**
         * Represents a collection of PivotTables associated with the workbook.
         */
        getPivotTables(): PivotTable[];

        /**
         * Add a PivotTable based on the specified source data and insert it at the top-left cell of the destination range.
         * @param name The name of the new PivotTable.
         * @param source The source data for the new PivotTable, this can either be a range (or string address including the worksheet name) or a table.
         * @param destination The cell in the upper-left corner of the PivotTable report's destination range (the range on the worksheet where the resulting report will be placed).
         */
        addPivotTable(
            name: string,
            source: Range | string | Table,
            destination: Range | string
        ): PivotTable;

        /**
         * Gets a PivotTable by name. If the PivotTable does not exist, then this method returns `undefined`.
         * @param name Name of the PivotTable to be retrieved.
         */
        getPivotTable(name: string): PivotTable | undefined;

        /**
         * Refreshes all the pivot tables in the collection.
         */
        refreshAllPivotTables(): void;

        /**
         * Represents a collection of SlicerStyles associated with the workbook.
         */
        getSlicerStyles(): SlicerStyle[];

        /**
         * Creates a blank slicer style with the specified name.
         * @param name The unique name for the new slicer style. Will throw an `InvalidArgument` exception if the name is already in use.
         * @param makeUniqueName Optional. Defaults to `false`. If `true`, will append numbers to the name in order to make it unique, if needed.
         */
        addSlicerStyle(name: string, makeUniqueName?: boolean): SlicerStyle;

        /**
         * Gets the default `SlicerStyle` for the parent object's scope.
         */
        getDefaultSlicerStyle(): SlicerStyle;

        /**
         * Gets a `SlicerStyle` by name. If the slicer style doesn't exist, then this method returns `undefined`.
         * @param name Name of the slicer style to be retrieved.
         */
        getSlicerStyle(name: string): SlicerStyle | undefined;

        /**
         * Sets the default slicer style for use in the parent object's scope.
         * @param newDefaultStyle The `SlicerStyle` object, or name of the `SlicerStyle` object, that should be the new default.
         */
        setDefaultSlicerStyle(newDefaultStyle: SlicerStyle | string): void;

        /**
         * Represents a collection of slicers associated with the workbook.
         */
        getSlicers(): Slicer[];

        /**
         * Adds a new slicer to the workbook.
         * @param slicerSource The data source that the new slicer will be based on. It can be a `PivotTable` object, a `Table` object, or a string. When a PivotTable object is passed, the data source is the source of the `PivotTable` object. When a `Table` object is passed, the data source is the `Table` object. When a string is passed, it is interpreted as the name or ID of a PivotTable or table.
         * @param sourceField The field in the data source to filter by. It can be a `PivotField` object, a `TableColumn` object, the ID of a `PivotField` or the name or ID of a `TableColumn`.
         * @param slicerDestination Optional. The worksheet in which the new slicer will be created. It can be a `Worksheet` object or the name or ID of a worksheet. This parameter can be omitted if the slicer collection is retrieved from a worksheet.
         */
        addSlicer(
            slicerSource: string | PivotTable | Table,
            sourceField: string | PivotField | number | TableColumn,
            slicerDestination?: string | Worksheet
        ): Slicer;

        /**
         * Gets a slicer using its name or ID. If the slicer doesn't exist, then this method returns `undefined`.
         * @param key Name or ID of the slicer to be retrieved.
         */
        getSlicer(key: string): Slicer | undefined;

        /**
         * Represents a collection of styles associated with the workbook.
         */
        getPredefinedCellStyles(): PredefinedCellStyle[];

        /**
         * Adds a new style to the collection.
         * @param name Name of the style to be added.
         */
        addPredefinedCellStyle(name: string): void;

        /**
         * Gets a `Style` by name.
         * @param name Name of the style to be retrieved.
         */
        getPredefinedCellStyle(name: string): PredefinedCellStyle;

        /**
         * Represents a collection of TableStyles associated with the workbook.
         */
        getTableStyles(): TableStyle[];

        /**
         * Creates a blank `TableStyle` with the specified name.
         * @param name The unique name for the new table style. Will throw an `InvalidArgument` error if the name is already in use.
         * @param makeUniqueName Optional. Defaults to `false`. If `true`, will append numbers to the name in order to make it unique, if needed.
         */
        addTableStyle(name: string, makeUniqueName?: boolean): TableStyle;

        /**
         * Gets the default table style for the parent object's scope.
         */
        getDefaultTableStyle(): TableStyle;

        /**
         * Gets a `TableStyle` by name. If the table style does not exist, then this method returns `undefined`.
         * @param name Name of the table style to be retrieved.
         */
        getTableStyle(name: string): TableStyle | undefined;

        /**
         * Sets the default table style for use in the parent object's scope.
         * @param newDefaultStyle The `TableStyle` object, or name of the `TableStyle` object, that should be the new default.
         */
        setDefaultTableStyle(newDefaultStyle: TableStyle | string): void;

        /**
         * Represents a collection of tables associated with the workbook.
         */
        getTables(): Table[];

        /**
         * Creates a new table. The range object or source address determines the worksheet under which the table will be added. If the table cannot be added (e.g., because the address is invalid, or the table would overlap with another table), an error will be thrown.
         * @param address A `Range` object, or a string address or name of the range representing the data source. If the address does not contain a sheet name, the currently-active sheet is used.
         * @param hasHeaders A boolean value that indicates whether the data being imported has column labels. If the source does not contain headers (i.e., when this property set to `false`), Excel will automatically generate a header and shift the data down by one row.
         */
        addTable(address: Range | string, hasHeaders: boolean): Table;

        /**
         * Gets a table by name or ID. If the table doesn't exist, then this method returns `undefined`.
         * @param key Name or ID of the table to be retrieved.
         */
        getTable(key: string): Table | undefined;

        /**
         * Represents a collection of TimelineStyles associated with the workbook.
         */
        getTimelineStyles(): TimelineStyle[];

        /**
         * Creates a blank `TimelineStyle` with the specified name.
         * @param name The unique name for the new timeline style. Will throw an `InvalidArgument` error if the name is already in use.
         * @param makeUniqueName Optional. Defaults to `false`. If `true`, will append numbers to the name in order to make it unique, if needed.
         */
        addTimelineStyle(name: string, makeUniqueName?: boolean): TimelineStyle;

        /**
         * Gets the default timeline style for the parent object's scope.
         */
        getDefaultTimelineStyle(): TimelineStyle;

        /**
         * Gets a `TimelineStyle` by name. If the timeline style doesn't exist, then this method returns `undefined`.
         * @param name Name of the timeline style to be retrieved.
         */
        getTimelineStyle(name: string): TimelineStyle | undefined;

        /**
         * Sets the default timeline style for use in the parent object's scope.
         * @param newDefaultStyle The `TimelineStyle` object, or name of the `TimelineStyle` object, that should be the new default.
         */
        setDefaultTimelineStyle(newDefaultStyle: TimelineStyle | string): void;

        /**
         * Represents a collection of worksheets associated with the workbook.
         */
        getWorksheets(): Worksheet[];

        /**
         * Adds a new worksheet to the workbook. The worksheet will be added at the end of existing worksheets. If you wish to activate the newly added worksheet, call `.activate()` on it.
         * @param name Optional. The name of the worksheet to be added. If specified, the name should be unique. If not specified, Excel determines the name of the new worksheet.
         */
        addWorksheet(name?: string): Worksheet;

        /**
         * Gets the currently active worksheet in the workbook.
         */
        getActiveWorksheet(): Worksheet;

        /**
         * Gets the first worksheet in the collection.
         * @param visibleOnly Optional. If `true`, considers only visible worksheets, skipping over any hidden ones.
         */
        getFirstWorksheet(visibleOnly?: boolean): Worksheet;

        /**
         * Gets a worksheet object using its name or ID. If the worksheet does not exist, then this method returns `undefined`.
         * @param key The name or ID of the worksheet.
         */
        getWorksheet(key: string): Worksheet | undefined;

        /**
         * Gets the last worksheet in the collection.
         * @param visibleOnly Optional. If `true`, considers only visible worksheets, skipping over any hidden ones.
         */
        getLastWorksheet(visibleOnly?: boolean): Worksheet;

        /**
         * Refreshes all the Data Connections.
         */
        refreshAllDataConnections(): void;

        /**
         * Gets a new collection of custom XML parts whose namespaces match the given namespace.
         * @param namespaceUri This must be a fully qualified schema URI; for example, "http://schemas.contoso.com/review/1.0".
         * @deprecated Use `getCustomXmlPartsByNamespace` instead.
         */
        getCustomXmlPartByNamespace(namespaceUri: string): CustomXmlPart[];
    }

    /**
     * Represents the protection of a workbook object.
     */
    interface WorkbookProtection {
        /**
         * Specifies if the workbook is protected.
         */
        getProtected(): boolean;

        /**
         * Protects a workbook. Fails if the workbook has been protected.
         * @param password Workbook protection password.
         */
        protect(password?: string): void;

        /**
         * Unprotects a workbook.
         * @param password Workbook protection password.
         */
        unprotect(password?: string): void;
    }

    /**
     * An Excel worksheet is a grid of cells. It can contain data, tables, charts, etc.
     */
    interface Worksheet {
        /**
         * Represents the `AutoFilter` object of the worksheet.
         */
        getAutoFilter(): AutoFilter;

        /**
         * Determines if Excel should recalculate the worksheet when necessary.
         * True if Excel recalculates the worksheet when necessary. False if Excel doesn't recalculate the sheet.
         */
        getEnableCalculation(): boolean;

        /**
         * Determines if Excel should recalculate the worksheet when necessary.
         * True if Excel recalculates the worksheet when necessary. False if Excel doesn't recalculate the sheet.
         */
        setEnableCalculation(enableCalculation: boolean): void;

        /**
         * Gets an object that can be used to manipulate frozen panes on the worksheet.
         */
        getFreezePanes(): WorksheetFreezePanes;

        /**
         * Returns a value that uniquely identifies the worksheet in a given workbook. The value of the identifier remains the same even when the worksheet is renamed or moved.
         */
        getId(): string;

        /**
         * The display name of the worksheet.
         */
        getName(): string;

        /**
         * The display name of the worksheet.
         */
        setName(name: string): void;

        /**
         * Gets the `PageLayout` object of the worksheet.
         */
        getPageLayout(): PageLayout;

        /**
         * The zero-based position of the worksheet within the workbook.
         */
        getPosition(): number;

        /**
         * The zero-based position of the worksheet within the workbook.
         */
        setPosition(position: number): void;

        /**
         * Returns the sheet protection object for a worksheet.
         */
        getProtection(): WorksheetProtection;

        /**
         * Specifies if gridlines are visible to the user.
         */
        getShowGridlines(): boolean;

        /**
         * Specifies if gridlines are visible to the user.
         */
        setShowGridlines(showGridlines: boolean): void;

        /**
         * Specifies if headings are visible to the user.
         */
        getShowHeadings(): boolean;

        /**
         * Specifies if headings are visible to the user.
         */
        setShowHeadings(showHeadings: boolean): void;

        /**
         * Returns the standard (default) height of all the rows in the worksheet, in points.
         */
        getStandardHeight(): number;

        /**
         * Specifies the standard (default) width of all the columns in the worksheet.
         * One unit of column width is equal to the width of one character in the Normal style. For proportional fonts, the width of the character 0 (zero) is used.
         */
        getStandardWidth(): number;

        /**
         * Specifies the standard (default) width of all the columns in the worksheet.
         * One unit of column width is equal to the width of one character in the Normal style. For proportional fonts, the width of the character 0 (zero) is used.
         */
        setStandardWidth(standardWidth: number): void;

        /**
         * The tab color of the worksheet.
         * When retrieving the tab color, if the worksheet is invisible, the value will be `null`. If the worksheet is visible but the tab color is set to auto, an empty string will be returned. Otherwise, the property will be set to a color, in the form #RRGGBB (e.g., "FFA500").
         * When setting the color, use an empty-string to set an "auto" color, or a real color otherwise.
         */
        getTabColor(): string;

        /**
         * The tab color of the worksheet.
         * When retrieving the tab color, if the worksheet is invisible, the value will be `null`. If the worksheet is visible but the tab color is set to auto, an empty string will be returned. Otherwise, the property will be set to a color, in the form #RRGGBB (e.g., "FFA500").
         * When setting the color, use an empty-string to set an "auto" color, or a real color otherwise.
         */
        setTabColor(tabColor: string): void;

        /**
         * The visibility of the worksheet.
         */
        getVisibility(): SheetVisibility;

        /**
         * The visibility of the worksheet.
         */
        setVisibility(visibility: SheetVisibility): void;

        /**
         * Activate the worksheet in the Excel UI.
         */
        activate(): void;

        /**
         * Calculates all cells on a worksheet.
         * @param markAllDirty True, to mark all as dirty.
         */
        calculate(markAllDirty: boolean): void;

        /**
         * Copies a worksheet and places it at the specified position.
         * @param positionType The location in the workbook to place the newly created worksheet. The default value is "None", which inserts the worksheet at the beginning of the worksheet.
         * @param relativeTo The existing worksheet which determines the newly created worksheet's position. This is only needed if `positionType` is "Before" or "After".
         */
        copy(
            positionType?: WorksheetPositionType,
            relativeTo?: Worksheet
        ): Worksheet;

        /**
         * Deletes the worksheet from the workbook. Note that if the worksheet's visibility is set to "VeryHidden", the delete operation will fail with an `InvalidOperation` exception. You should first change its visibility to hidden or visible before deleting it.
         */
        delete(): void;

        /**
         * Finds all occurrences of the given string based on the criteria specified and returns them as a `RangeAreas` object, comprising one or more rectangular ranges.
         * @param text The string to find.
         * @param criteria Additional search criteria, including whether the search needs to match the entire cell or be case-sensitive.
         */
        findAll(text: string, criteria: WorksheetSearchCriteria): RangeAreas;

        /**
         * Gets the `Range` object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range, so long as it stays within the worksheet grid.
         * @param row The row number of the cell to be retrieved. Zero-indexed.
         * @param column The column number of the cell to be retrieved. Zero-indexed.
         */
        getCell(row: number, column: number): Range;

        /**
         * Gets the worksheet that follows this one. If there are no worksheets following this one, then this method returns `undefined`.
         * @param visibleOnly Optional. If `true`, considers only visible worksheets, skipping over any hidden ones.
         */
        getNext(visibleOnly?: boolean): Worksheet;

        /**
         * Gets the worksheet that precedes this one. If there are no previous worksheets, then this method returns `undefined`.
         * @param visibleOnly Optional. If `true`, considers only visible worksheets, skipping over any hidden ones.
         */
        getPrevious(visibleOnly?: boolean): Worksheet;

        /**
         * Gets the `Range` object, representing a single rectangular block of cells, specified by the address or name.
         * @param address Optional. The string representing the address or name of the range. For example, "A1:B2". If not specified, the entire worksheet range is returned.
         */
        getRange(address?: string): Range;

        /**
         * Gets the `Range` object beginning at a particular row index and column index, and spanning a certain number of rows and columns.
         * @param startRow Start row (zero-indexed).
         * @param startColumn Start column (zero-indexed).
         * @param rowCount Number of rows to include in the range.
         * @param columnCount Number of columns to include in the range.
         */
        getRangeByIndexes(
            startRow: number,
            startColumn: number,
            rowCount: number,
            columnCount: number
        ): Range;

        /**
         * Gets the `RangeAreas` object, representing one or more blocks of rectangular ranges, specified by the address or name.
         * @param address Optional. A string containing the comma-separated or semicolon-separated addresses or names of the individual ranges. For example, "A1:B2, A5:B5" or "A1:B2; A5:B5". If not specified, a `RangeAreas` object for the entire worksheet is returned.
         */
        getRanges(address?: string): RangeAreas;

        /**
         * @param valuesOnly Optional. Considers only cells with values as used cells.
         */
        getUsedRange(valuesOnly?: boolean): Range;

        /**
         * Finds and replaces the given string based on the criteria specified within the current worksheet.
         * @param text String to find.
         * @param replacement The string that replaces the original string.
         * @param criteria Additional replacement criteria.
         */
        replaceAll(
            text: string,
            replacement: string,
            criteria: ReplaceCriteria
        ): number;

        /**
         * Shows row or column groups by their outline levels.
         * Outlines groups and summarizes a list of data in the worksheet.
         * The `rowLevels` and `columnLevels` parameters specify how many levels of the outline will be displayed.
         * The acceptable argument range is between 0 and 8.
         * A value of 0 does not change the current display. A value greater than the current number of levels displays all the levels.
         * @param rowLevels The number of row levels of an outline to display.
         * @param columnLevels The number of column levels of an outline to display.
         */
        showOutlineLevels(rowLevels: number, columnLevels: number): void;

        /**
         * Returns a collection of charts that are part of the worksheet.
         */
        getCharts(): Chart[];

        /**
         * Creates a new chart.
         * @param type Represents the type of a chart. See `ExcelScript.ChartType` for details.
         * @param sourceData The `Range` object corresponding to the source data.
         * @param seriesBy Optional. Specifies the way columns or rows are used as data series on the chart. See `ExcelScript.ChartSeriesBy` for details.
         */
        addChart(
            type: ChartType,
            sourceData: Range,
            seriesBy?: ChartSeriesBy
        ): Chart;

        /**
         * Gets a chart using its name. If there are multiple charts with the same name, the first one will be returned. If the chart doesn't exist, then this method returns `undefined`.
         * @param name Name of the chart to be retrieved.
         */
        getChart(name: string): Chart | undefined;

        /**
         * Returns a collection of all the Comments objects on the worksheet.
         */
        getComments(): Comment[];

        /**
         * Creates a new comment with the given content on the given cell. An `InvalidArgument` error is thrown if the provided range is larger than one cell.
         * @param cellAddress The cell to which the comment is added. This can be a `Range` object or a string. If it's a string, it must contain the full address, including the sheet name. An `InvalidArgument` error is thrown if the provided range is larger than one cell.
         * @param content The comment's content. This can be either a string or `CommentRichContent` object. Strings are used for plain text. `CommentRichContent` objects allow for other comment features, such as mentions.
         * @param contentType Optional. The type of content contained within the comment. The default value is enum `ContentType.Plain`.
         */
        addComment(
            cellAddress: Range | string,
            content: CommentRichContent | string,
            contentType?: ContentType
        ): Comment;

        /**
         * Gets a comment from the collection based on its ID.
         * @param commentId The identifier for the comment.
         */
        getComment(commentId: string): Comment;

        /**
         * Gets the comment from the specified cell.
         * @param cellAddress The cell which the comment is on. This can be a `Range` object or a string. If it's a string, it must contain the full address, including the sheet name. An `InvalidArgument` error is thrown if the provided range is larger than one cell.
         */
        getCommentByCell(cellAddress: Range | string): Comment;

        /**
         * Gets the comment to which the given reply is connected.
         * @param replyId The identifier of comment reply.
         */
        getCommentByReplyId(replyId: string): Comment;

        /**
         * Gets a collection of worksheet-level custom properties.
         */
        getCustomProperties(): WorksheetCustomProperty[];

        /**
         * Adds a new custom property that maps to the provided key. This overwrites existing custom properties with that key.
         * @param key The key that identifies the custom property object. It is case-insensitive.The key is limited to 255 characters (larger values will cause an `InvalidArgument` error to be thrown.)
         * @param value The value of this custom property.
         */
        addWorksheetCustomProperty(
            key: string,
            value: string
        ): WorksheetCustomProperty;

        /**
         * Gets a custom property object by its key, which is case-insensitive. If the custom property doesn't exist, then this method returns `undefined`.
         * @param key The key that identifies the custom property object. It is case-insensitive.
         */
        getWorksheetCustomProperty(
            key: string
        ): WorksheetCustomProperty | undefined;

        /**
         * Gets the horizontal page break collection for the worksheet. This collection only contains manual page breaks.
         */
        getHorizontalPageBreaks(): PageBreak[];

        /**
         * Adds a page break before the top-left cell of the range specified.
         * @param pageBreakRange The range immediately after the page break to be added.
         */
        addHorizontalPageBreak(pageBreakRange: Range | string): PageBreak;

        /**
         * Resets all manual page breaks in the collection.
         */
        removeAllHorizontalPageBreaks(): void;

        /**
         * Returns a collection of sheet views that are present in the worksheet.
         */
        getNamedSheetViews(): NamedSheetView[];

        /**
         * Creates a new sheet view with the given name.
         * @param name The name of the sheet view to be created.
         * Throws an error when the provided name already exists, is empty, or is a name reserved by the worksheet.
         */
        addNamedSheetView(name: string): NamedSheetView;

        /**
         * Creates and activates a new temporary sheet view.
         * Temporary views are removed when closing the application, exiting the temporary view with the exit method, or switching to another sheet view.
         * The temporary sheet view can also be accessed with the empty string (""), if the temporary view exists.
         */
        enterTemporaryNamedSheetView(): NamedSheetView;

        /**
         * Exits the currently active sheet view.
         */
        exitActiveNamedSheetView(): void;

        /**
         * Gets the worksheet's currently active sheet view.
         */
        getActiveNamedSheetView(): NamedSheetView;

        /**
         * Gets a sheet view using its name.
         * @param key The case-sensitive name of the sheet view. Use the empty string ("") to get the temporary sheet view, if the temporary view exists.
         */
        getNamedSheetView(key: string): NamedSheetView;

        /**
         * Collection of names scoped to the current worksheet.
         */
        getNames(): NamedItem[];

        /**
         * Adds a new name to the collection of the given scope.
         * @param name The name of the named item.
         * @param reference The formula or the range that the name will refer to.
         * @param comment Optional. The comment associated with the named item.
         */
        addNamedItem(
            name: string,
            reference: Range | string,
            comment?: string
        ): NamedItem;

        /**
         * Adds a new name to the collection of the given scope using the user's locale for the formula.
         * @param name The name of the named item.
         * @param formula The formula in the user's locale that the name will refer to.
         * @param comment Optional. The comment associated with the named item.
         */
        addNamedItemFormulaLocal(
            name: string,
            formula: string,
            comment?: string
        ): NamedItem;

        /**
         * Gets a `NamedItem` object using its name. If the object does not exist, then this method returns `undefined`.
         * @param name Nameditem name.
         */
        getNamedItem(name: string): NamedItem | undefined;

        /**
         * Collection of PivotTables that are part of the worksheet.
         */
        getPivotTables(): PivotTable[];

        /**
         * Add a PivotTable based on the specified source data and insert it at the top-left cell of the destination range.
         * @param name The name of the new PivotTable.
         * @param source The source data for the new PivotTable, this can either be a range (or string address including the worksheet name) or a table.
         * @param destination The cell in the upper-left corner of the PivotTable report's destination range (the range on the worksheet where the resulting report will be placed).
         */
        addPivotTable(
            name: string,
            source: Range | string | Table,
            destination: Range | string
        ): PivotTable;

        /**
         * Gets a PivotTable by name. If the PivotTable does not exist, then this method returns `undefined`.
         * @param name Name of the PivotTable to be retrieved.
         */
        getPivotTable(name: string): PivotTable | undefined;

        /**
         * Refreshes all the pivot tables in the collection.
         */
        refreshAllPivotTables(): void;

        /**
         * Returns the collection of all the Shape objects on the worksheet.
         */
        getShapes(): Shape[];

        /**
         * Adds a geometric shape to the worksheet. Returns a `Shape` object that represents the new shape.
         * @param geometricShapeType Represents the type of the geometric shape. See `ExcelScript.GeometricShapeType` for details.
         */
        addGeometricShape(geometricShapeType: GeometricShapeType): Shape;

        /**
         * Groups a subset of shapes in this collection's worksheet. Returns a `Shape` object that represents the new group of shapes.
         * @param values An array of shape IDs or shape objects.
         */
        addGroup(values: Array<string | Shape>): Shape;

        /**
         * Creates an image from a base64-encoded string and adds it to the worksheet. Returns the `Shape` object that represents the new image.
         * @param base64ImageString A base64-encoded string representing an image in either JPEG or PNG format.
         */
        addImage(base64ImageString: string): Shape;

        /**
         * Adds a line to worksheet. Returns a `Shape` object that represents the new line.
         * @param startLeft The distance, in points, from the start of the line to the left side of the worksheet.
         * @param startTop The distance, in points, from the start of the line to the top of the worksheet.
         * @param endLeft The distance, in points, from the end of the line to the left of the worksheet.
         * @param endTop The distance, in points, from the end of the line to the top of the worksheet.
         * @param connectorType Represents the connector type. See `ExcelScript.ConnectorType` for details.
         */
        addLine(
            startLeft: number,
            startTop: number,
            endLeft: number,
            endTop: number,
            connectorType?: ConnectorType
        ): Shape;

        /**
         * Adds a text box to the worksheet with the provided text as the content. Returns a `Shape` object that represents the new text box.
         * @param text Represents the text that will be shown in the created text box.
         */
        addTextBox(text?: string): Shape;

        /**
         * Gets a shape using its name or ID.
         * @param key The name or ID of the shape to be retrieved.
         */
        getShape(key: string): Shape;

        /**
         * Returns a collection of slicers that are part of the worksheet.
         */
        getSlicers(): Slicer[];

        /**
         * Adds a new slicer to the workbook.
         * @param slicerSource The data source that the new slicer will be based on. It can be a `PivotTable` object, a `Table` object, or a string. When a PivotTable object is passed, the data source is the source of the `PivotTable` object. When a `Table` object is passed, the data source is the `Table` object. When a string is passed, it is interpreted as the name or ID of a PivotTable or table.
         * @param sourceField The field in the data source to filter by. It can be a `PivotField` object, a `TableColumn` object, the ID of a `PivotField` or the name or ID of a `TableColumn`.
         * @param slicerDestination Optional. The worksheet in which the new slicer will be created. It can be a `Worksheet` object or the name or ID of a worksheet. This parameter can be omitted if the slicer collection is retrieved from a worksheet.
         */
        addSlicer(
            slicerSource: string | PivotTable | Table,
            sourceField: string | PivotField | number | TableColumn,
            slicerDestination?: string | Worksheet
        ): Slicer;

        /**
         * Gets a slicer using its name or ID. If the slicer doesn't exist, then this method returns `undefined`.
         * @param key Name or ID of the slicer to be retrieved.
         */
        getSlicer(key: string): Slicer | undefined;

        /**
         * Collection of tables that are part of the worksheet.
         */
        getTables(): Table[];

        /**
         * Creates a new table. The range object or source address determines the worksheet under which the table will be added. If the table cannot be added (e.g., because the address is invalid, or the table would overlap with another table), an error will be thrown.
         * @param address A `Range` object, or a string address or name of the range representing the data source. If the address does not contain a sheet name, the currently-active sheet is used.
         * @param hasHeaders A boolean value that indicates whether the data being imported has column labels. If the source does not contain headers (i.e., when this property set to `false`), Excel will automatically generate a header and shift the data down by one row.
         */
        addTable(address: Range | string, hasHeaders: boolean): Table;

        /**
         * Gets a table by name or ID. If the table doesn't exist, then this method returns `undefined`.
         * @param key Name or ID of the table to be retrieved.
         */
        getTable(key: string): Table | undefined;

        /**
         * Gets the vertical page break collection for the worksheet. This collection only contains manual page breaks.
         */
        getVerticalPageBreaks(): PageBreak[];

        /**
         * Adds a page break before the top-left cell of the range specified.
         * @param pageBreakRange The range immediately after the page break to be added.
         */
        addVerticalPageBreak(pageBreakRange: Range | string): PageBreak;

        /**
         * Resets all manual page breaks in the collection.
         */
        removeAllVerticalPageBreaks(): void;
    }

    /**
     * Represents the protection of a worksheet object.
     */
    interface WorksheetProtection {
        /**
         * Specifies the protection options for the worksheet.
         */
        getOptions(): WorksheetProtectionOptions;

        /**
         * Specifies if the worksheet is protected.
         */
        getProtected(): boolean;

        /**
         * Protects a worksheet. Fails if the worksheet has already been protected.
         * @param options Optional. Sheet protection options.
         * @param password Optional. Sheet protection password.
         */
        protect(options?: WorksheetProtectionOptions, password?: string): void;

        /**
         * Unprotects a worksheet.
         * @param password Sheet protection password.
         */
        unprotect(password?: string): void;
    }

    interface WorksheetFreezePanes {
        /**
         * Sets the frozen cells in the active worksheet view.
         * The range provided corresponds to cells that will be frozen in the top- and left-most pane.
         * @param frozenRange A range that represents the cells to be frozen, or `null` to remove all frozen panes.
         */
        freezeAt(frozenRange: Range | string): void;

        /**
         * Freeze the first column or columns of the worksheet in place.
         * @param count Optional number of columns to freeze, or zero to unfreeze all columns
         */
        freezeColumns(count?: number): void;

        /**
         * Freeze the top row or rows of the worksheet in place.
         * @param count Optional number of rows to freeze, or zero to unfreeze all rows
         */
        freezeRows(count?: number): void;

        /**
         * Gets a range that describes the frozen cells in the active worksheet view.
         * The frozen range corresponds to cells that are frozen in the top- and left-most pane.
         * If there is no frozen pane, then this method returns `undefined`.
         */
        getLocation(): Range;

        /**
         * Removes all frozen panes in the worksheet.
         */
        unfreeze(): void;
    }

    /**
     * Range represents a set of one or more contiguous cells such as a cell, a row, a column, or a block of cells.
     */
    interface Range {
        /**
         * Specifies the range reference in A1-style. Address value contains the sheet reference (e.g., "Sheet1!A1:B4").
         */
        getAddress(): string;

        /**
         * Represents the range reference for the specified range in the language of the user.
         */
        getAddressLocal(): string;

        /**
         * Specifies the number of cells in the range. This API will return -1 if the cell count exceeds 2^31-1 (2,147,483,647).
         */
        getCellCount(): number;

        /**
         * Specifies the total number of columns in the range.
         */
        getColumnCount(): number;

        /**
         * Represents if all columns in the current range are hidden. Value is `true` when all columns in a range are hidden. Value is `false` when no columns in the range are hidden. Value is `null` when some columns in a range are hidden and other columns in the same range are not hidden.
         */
        getColumnHidden(): boolean;

        /**
         * Represents if all columns in the current range are hidden. Value is `true` when all columns in a range are hidden. Value is `false` when no columns in the range are hidden. Value is `null` when some columns in a range are hidden and other columns in the same range are not hidden.
         */
        setColumnHidden(columnHidden: boolean): void;

        /**
         * Specifies the column number of the first cell in the range. Zero-indexed.
         */
        getColumnIndex(): number;

        /**
         * Returns a data validation object.
         */
        getDataValidation(): DataValidation;

        /**
         * Returns a format object, encapsulating the range's font, fill, borders, alignment, and other properties.
         */
        getFormat(): RangeFormat;

        /**
         * Represents the formula in A1-style notation. If a cell has no formula, its value is returned instead.
         */
        getFormulas(): string[][];

        /**
         * Represents the formula in A1-style notation. If a cell has no formula, its value is returned instead.
         */
        setFormulas(formulas: string[][]): void;

        /**
         * Represents the formula in A1-style notation, in the user's language and number-formatting locale. For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German. If a cell has no formula, its value is returned instead.
         */
        getFormulasLocal(): string[][];

        /**
         * Represents the formula in A1-style notation, in the user's language and number-formatting locale. For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German. If a cell has no formula, its value is returned instead.
         */
        setFormulasLocal(formulasLocal: string[][]): void;

        /**
         * Represents the formula in R1C1-style notation. If a cell has no formula, its value is returned instead.
         */
        getFormulasR1C1(): string[][];

        /**
         * Represents the formula in R1C1-style notation. If a cell has no formula, its value is returned instead.
         */
        setFormulasR1C1(formulasR1C1: string[][]): void;

        /**
         * Represents if all cells have a spill border.
         * Returns `true` if all cells have a spill border, or `false` if all cells do not have a spill border.
         * Returns `null` if there are cells both with and without spill borders within the range.
         */
        getHasSpill(): boolean;

        /**
         * Returns the distance in points, for 100% zoom, from the top edge of the range to the bottom edge of the range.
         */
        getHeight(): number;

        /**
         * Represents if all cells in the current range are hidden. Value is `true` when all cells in a range are hidden. Value is `false` when no cells in the range are hidden. Value is `null` when some cells in a range are hidden and other cells in the same range are not hidden.
         */
        getHidden(): boolean;

        /**
         * Represents the hyperlink for the current range.
         */
        getHyperlink(): RangeHyperlink;

        /**
         * Represents the hyperlink for the current range.
         */
        setHyperlink(hyperlink: RangeHyperlink): void;

        /**
         * Represents if the current range is an entire column.
         */
        getIsEntireColumn(): boolean;

        /**
         * Represents if the current range is an entire row.
         */
        getIsEntireRow(): boolean;

        /**
         * Returns the distance in points, for 100% zoom, from the left edge of the worksheet to the left edge of the range.
         */
        getLeft(): number;

        /**
         * Represents the data type state of each cell.
         */
        getLinkedDataTypeStates(): LinkedDataTypeState[][];

        /**
         * Represents Excel's number format code for the given range.
         */
        getNumberFormats(): string[][];

        /**
         * Represents Excel's number format code for the given range.
         */
        setNumberFormats(numberFormats: string[][]): void;

        /**
         * Represents the category of number format of each cell.
         */
        getNumberFormatCategories(): NumberFormatCategory[][];

        /**
         * Represents Excel's number format code for the given range, based on the language settings of the user.
         * Excel does not perform any language or format coercion when getting or setting the `numberFormatLocal` property.
         * Any returned text uses the locally-formatted strings based on the language specified in the system settings.
         */
        getNumberFormatsLocal(): string[][];

        /**
         * Represents Excel's number format code for the given range, based on the language settings of the user.
         * Excel does not perform any language or format coercion when getting or setting the `numberFormatLocal` property.
         * Any returned text uses the locally-formatted strings based on the language specified in the system settings.
         */
        setNumberFormatsLocal(numberFormatsLocal: string[][]): void;

        /**
         * Returns the total number of rows in the range.
         */
        getRowCount(): number;

        /**
         * Represents if all rows in the current range are hidden. Value is `true` when all rows in a range are hidden. Value is `false` when no rows in the range are hidden. Value is `null` when some rows in a range are hidden and other rows in the same range are not hidden.
         */
        getRowHidden(): boolean;

        /**
         * Represents if all rows in the current range are hidden. Value is `true` when all rows in a range are hidden. Value is `false` when no rows in the range are hidden. Value is `null` when some rows in a range are hidden and other rows in the same range are not hidden.
         */
        setRowHidden(rowHidden: boolean): void;

        /**
         * Returns the row number of the first cell in the range. Zero-indexed.
         */
        getRowIndex(): number;

        /**
         * Represents if all the cells would be saved as an array formula.
         * Returns `true` if all cells would be saved as an array formula, or `false` if all cells would not be saved as an array formula.
         * Returns `null` if some cells would be saved as an array formula and some would not be.
         */
        getSavedAsArray(): boolean;

        /**
         * Represents the range sort of the current range.
         */
        getSort(): RangeSort;

        /**
         * Represents the style of the current range.
         * If the styles of the cells are inconsistent, `null` will be returned.
         * For custom styles, the style name will be returned. For built-in styles, a string representing a value in the `BuiltInStyle` enum will be returned.
         */
        getPredefinedCellStyle(): string;

        /**
         * Represents the style of the current range.
         * If the styles of the cells are inconsistent, `null` will be returned.
         * For custom styles, the style name will be returned. For built-in styles, a string representing a value in the `BuiltInStyle` enum will be returned.
         */
        setPredefinedCellStyle(predefinedCellStyle: string): void;

        /**
         * Text values of the specified range. The text value will not depend on the cell width. The number sign (#) substitution that happens in the Excel UI will not affect the text value returned by the API.
         */
        getTexts(): string[][];

        /**
         * Returns the distance in points, for 100% zoom, from the top edge of the worksheet to the top edge of the range.
         */
        getTop(): number;

        /**
         * Specifies the type of data in each cell.
         */
        getValueTypes(): RangeValueType[][];

        /**
         * Represents the raw values of the specified range. The data returned could be a string, number, or boolean. Cells that contain an error will return the error string.
         * If the returned value starts with a plus ("+"), minus ("-"), or equal sign ("="), Excel interprets this value as a formula.
         */
        getValues(): (string | number | boolean)[][];

        /**
         * Sets the raw values of the specified range. The data provided could be a string, number, or boolean.
         * If the provided value starts with a plus ("+"), minus ("-"), or equal sign ("="), Excel interprets this value as a formula.
         */
        setValues(values: (string | number | boolean)[][]): void;

        /**
         * Returns the distance in points, for 100% zoom, from the left edge of the range to the right edge of the range.
         */
        getWidth(): number;

        /**
         * The worksheet containing the current range.
         */
        getWorksheet(): Worksheet;

        /**
         * Fills a range from the current range to the destination range using the specified AutoFill logic.
         * The destination range can be `null` or can extend the source range either horizontally or vertically.
         * Discontiguous ranges are not supported.
         *
         * @param destinationRange The destination range to AutoFill. If the destination range is `null`, data is filled out based on the surrounding cells (which is the behavior when double-clicking the UI’s range fill handle).
         * @param autoFillType The type of AutoFill. Specifies how the destination range is to be filled, based on the contents of the current range. Default is "FillDefault".
         */
        autoFill(
            destinationRange?: Range | string,
            autoFillType?: AutoFillType
        ): void;

        /**
         * Calculates a range of cells on a worksheet.
         */
        calculate(): void;

        /**
         * Clear range values, format, fill, border, etc.
         * @param applyTo Optional. Determines the type of clear action. See `ExcelScript.ClearApplyTo` for details.
         */
        clear(applyTo?: ClearApplyTo): void;

        /**
         * Converts the range cells with data types into text.
         */
        convertDataTypeToText(): void;

        /**
         * Copies cell data or formatting from the source range or `RangeAreas` to the current range.
         * The destination range can be a different size than the source range or `RangeAreas`. The destination will be expanded automatically if it is smaller than the source.
         * @param sourceRange The source range or `RangeAreas` to copy from. When the source `RangeAreas` has multiple ranges, their form must be able to be created by removing full rows or columns from a rectangular range.
         * @param copyType The type of cell data or formatting to copy over. Default is "All".
         * @param skipBlanks True if to skip blank cells in the source range. Default is false.
         * @param transpose True if to transpose the cells in the destination range. Default is false.
         */
        copyFrom(
            sourceRange: Range | RangeAreas | string,
            copyType?: RangeCopyType,
            skipBlanks?: boolean,
            transpose?: boolean
        ): void;

        /**
         * Deletes the cells associated with the range.
         * @param shift Specifies which way to shift the cells. See `ExcelScript.DeleteShiftDirection` for details.
         */
        delete(shift: DeleteShiftDirection): void;

        /**
         * Finds the given string based on the criteria specified.
         * If the current range is larger than a single cell, then the search will be limited to that range, else the search will cover the entire sheet starting after that cell.
         * If there are no matches, then this method returns `undefined`.
         * @param text The string to find.
         * @param criteria Additional search criteria, including the search direction and whether the search needs to match the entire cell or be case-sensitive.
         */
        find(text: string, criteria: SearchCriteria): Range;

        /**
         * Does a Flash Fill to the current range. Flash Fill automatically fills data when it senses a pattern, so the range must be a single column range and have data around it in order to find a pattern.
         */
        flashFill(): void;

        /**
         * Gets a `Range` object with the same top-left cell as the current `Range` object, but with the specified numbers of rows and columns.
         * @param numRows The number of rows of the new range size.
         * @param numColumns The number of columns of the new range size.
         */
        getAbsoluteResizedRange(numRows: number, numColumns: number): Range;

        /**
         * Gets the smallest range object that encompasses the given ranges. For example, the `GetBoundingRect` of "B2:C5" and "D10:E15" is "B2:E15".
         * @param anotherRange The range object, address, or range name.
         */
        getBoundingRect(anotherRange: Range | string): Range;

        /**
         * Gets the range object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range, so long as it stays within the worksheet grid. The returned cell is located relative to the top left cell of the range.
         * @param row Row number of the cell to be retrieved. Zero-indexed.
         * @param column Column number of the cell to be retrieved. Zero-indexed.
         */
        getCell(row: number, column: number): Range;

        /**
         * Gets a column contained in the range.
         * @param column Column number of the range to be retrieved. Zero-indexed.
         */
        getColumn(column: number): Range;

        /**
         * Gets a certain number of columns to the right of the current `Range` object.
         * @param count Optional. The number of columns to include in the resulting range. In general, use a positive number to create a range outside the current range. You can also use a negative number to create a range within the current range. The default value is 1.
         */
        getColumnsAfter(count?: number): Range;

        /**
         * Gets a certain number of columns to the left of the current `Range` object.
         * @param count Optional. The number of columns to include in the resulting range. In general, use a positive number to create a range outside the current range. You can also use a negative number to create a range within the current range. The default value is 1.
         */
        getColumnsBefore(count?: number): Range;

        /**
         * Returns a `WorkbookRangeAreas` object that represents the range containing all the direct precedents of a cell in the same worksheet or in multiple worksheets.
         */
        getDirectPrecedents(): WorkbookRangeAreas;

        /**
         * Gets an object that represents the entire column of the range (for example, if the current range represents cells "B4:E11", its `getEntireColumn` is a range that represents columns "B:E").
         */
        getEntireColumn(): Range;

        /**
         * Gets an object that represents the entire row of the range (for example, if the current range represents cells "B4:E11", its `GetEntireRow` is a range that represents rows "4:11").
         */
        getEntireRow(): Range;

        /**
         * Returns a range object that includes the current range and up to the edge of the range, based on the provided direction. This matches the Ctrl+Shift+Arrow key behavior in the Excel on Windows UI.
         * @param direction The direction from the active cell.
         * @param activeCell The active cell in this range. By default, the active cell is the top-left cell of the range. An error is thrown if the active cell is not in this range.
         */
        getExtendedRange(
            direction: KeyboardDirection,
            activeCell?: Range | string
        ): Range;

        /**
         * Renders the range as a base64-encoded png image.
         * 
         * **Note**: There is a known issue with `Range.getImage` that causes wrapped text or text that exceeds the cell width to render on the same line without line wrapping. 
         * This makes the resulting image unreadable, since the text overflows across the entire row. 
         * As a workaround, make sure the data in the range fits in each of the cells as a single line.
         */
        getImage(): string;

        /**
         * Gets the range object that represents the rectangular intersection of the given ranges. If no intersection is found, then this method returns `undefined`.
         * @param anotherRange The range object or range address that will be used to determine the intersection of ranges.
         */
        getIntersection(anotherRange: Range | string): Range;

        /**
         * Gets the last cell within the range. For example, the last cell of "B2:D5" is "D5".
         */
        getLastCell(): Range;

        /**
         * Gets the last column within the range. For example, the last column of "B2:D5" is "D2:D5".
         */
        getLastColumn(): Range;

        /**
         * Gets the last row within the range. For example, the last row of "B2:D5" is "B5:D5".
         */
        getLastRow(): Range;

        /**
         * Returns a `RangeAreas` object that represents the merged areas in this range. Note that if the merged areas count in this range is more than 512, then this method will fail to return the result. If the `RangeAreas` object doesn't exist, then this method will return `undefined`.
         */
        getMergedAreas(): RangeAreas;

        /**
         * Gets an object which represents a range that's offset from the specified range. The dimension of the returned range will match this range. If the resulting range is forced outside the bounds of the worksheet grid, an error will be thrown.
         * @param rowOffset The number of rows (positive, negative, or 0) by which the range is to be offset. Positive values are offset downward, and negative values are offset upward.
         * @param columnOffset The number of columns (positive, negative, or 0) by which the range is to be offset. Positive values are offset to the right, and negative values are offset to the left.
         */
        getOffsetRange(rowOffset: number, columnOffset: number): Range;

        /**
         * Gets a scoped collection of PivotTables that overlap with the range.
         * @param fullyContained If `true`, returns only PivotTables that are fully contained within the range bounds. The default value is `false`.
         */
        getPivotTables(fullyContained?: boolean): PivotTable[];

        /**
         * Returns a range object that is the edge cell of the data region that corresponds to the provided direction. This matches the Ctrl+Arrow key behavior in the Excel on Windows UI.
         * @param direction The direction from the active cell.
         * @param activeCell The active cell in this range. By default, the active cell is the top-left cell of the range. An error is thrown if the active cell is not in this range.
         */
        getRangeEdge(
            direction: KeyboardDirection,
            activeCell?: Range | string
        ): Range;

        /**
         * Gets a `Range` object similar to the current `Range` object, but with its bottom-right corner expanded (or contracted) by some number of rows and columns.
         * @param deltaRows The number of rows by which to expand the bottom-right corner, relative to the current range. Use a positive number to expand the range, or a negative number to decrease it.
         * @param deltaColumns The number of columns by which to expand the bottom-right corner, relative to the current range. Use a positive number to expand the range, or a negative number to decrease it.
         */
        getResizedRange(deltaRows: number, deltaColumns: number): Range;

        /**
         * Gets a row contained in the range.
         * @param row Row number of the range to be retrieved. Zero-indexed.
         */
        getRow(row: number): Range;

        /**
         * Gets a certain number of rows above the current `Range` object.
         * @param count Optional. The number of rows to include in the resulting range. In general, use a positive number to create a range outside the current range. You can also use a negative number to create a range within the current range. The default value is 1.
         */
        getRowsAbove(count?: number): Range;

        /**
         * Gets a certain number of rows below the current `Range` object.
         * @param count Optional. The number of rows to include in the resulting range. In general, use a positive number to create a range outside the current range. You can also use a negative number to create a range within the current range. The default value is 1.
         */
        getRowsBelow(count?: number): Range;

        /**
         * Gets the `RangeAreas` object, comprising one or more ranges, that represents all the cells that match the specified type and value.
         * If no special cells are found, then this method returns `undefined`.
         * @param cellType The type of cells to include.
         * @param cellValueType If `cellType` is either `constants` or `formulas`, this argument is used to determine which types of cells to include in the result. These values can be combined together to return more than one type. The default is to select all constants or formulas, no matter what the type.
         */
        getSpecialCells(
            cellType: SpecialCellType,
            cellValueType?: SpecialCellValueType
        ): RangeAreas;

        /**
         * Gets the range object containing the anchor cell for the cell getting spilled into.
         * If it's not a spilled cell, or more than one cell is given, then this method returns `undefined`.
         */
        getSpillParent(): Range;

        /**
         * Gets the range object containing the spill range when called on an anchor cell.
         * If the range isn't an anchor cell or the spill range can't be found, then this method returns `undefined`.
         */
        getSpillingToRange(): Range;

        /**
         * Returns a `Range` object that represents the surrounding region for the top-left cell in this range. A surrounding region is a range bounded by any combination of blank rows and blank columns relative to this range.
         */
        getSurroundingRegion(): Range;

        /**
         * Gets a scoped collection of tables that overlap with the range.
         * @param fullyContained If `true`, returns only tables that are fully contained within the range bounds. The default value is `false`.
         */
        getTables(fullyContained?: boolean): Table[];

        /**
         * Returns the used range of the given range object. If there are no used cells within the range, then this method returns `undefined`.
         * @param valuesOnly Considers only cells with values as used cells.
         */
        getUsedRange(valuesOnly?: boolean): Range;

        /**
         * Represents the visible rows of the current range.
         */
        getVisibleView(): RangeView;

        /**
         * Groups columns and rows for an outline.
         * @param groupOption Specifies how the range can be grouped by rows or columns.
         * An `InvalidArgument` error is thrown when the group option differs from the range's
         * `isEntireRow` or `isEntireColumn` property (i.e., `range.isEntireRow` is true and `groupOption` is "ByColumns"
         * or `range.isEntireColumn` is true and `groupOption` is "ByRows").
         */
        group(groupOption: GroupOption): void;

        /**
         * Hides the details of the row or column group.
         * @param groupOption Specifies whether to hide the details of grouped rows or grouped columns.
         */
        hideGroupDetails(groupOption: GroupOption): void;

        /**
         * Inserts a cell or a range of cells into the worksheet in place of this range, and shifts the other cells to make space. Returns a new `Range` object at the now blank space.
         * @param shift Specifies which way to shift the cells. See `ExcelScript.InsertShiftDirection` for details.
         */
        insert(shift: InsertShiftDirection): Range;

        /**
         * Merge the range cells into one region in the worksheet.
         * @param across Optional. Set `true` to merge cells in each row of the specified range as separate merged cells. The default value is `false`.
         */
        merge(across?: boolean): void;

        /**
         * Moves cell values, formatting, and formulas from current range to the destination range, replacing the old information in those cells.
         * The destination range will be expanded automatically if it is smaller than the current range. Any cells in the destination range that are outside of the original range's area are not changed.
         * @param destinationRange destinationRange Specifies the range to where the information in this range will be moved.
         */
        moveTo(destinationRange: Range | string): void;

        /**
         * Removes duplicate values from the range specified by the columns.
         * @param columns The columns inside the range that may contain duplicates. At least one column needs to be specified. Zero-indexed.
         * @param includesHeader True if the input data contains header. Default is false.
         */
        removeDuplicates(
            columns: number[],
            includesHeader: boolean
        ): RemoveDuplicatesResult;

        /**
         * Finds and replaces the given string based on the criteria specified within the current range.
         * @param text String to find.
         * @param replacement The string that replaces the original string.
         * @param criteria Additional replacement criteria.
         */
        replaceAll(
            text: string,
            replacement: string,
            criteria: ReplaceCriteria
        ): number;

        /**
         * Selects the specified range in the Excel UI.
         */
        select(): void;

        /**
         * Set a range to be recalculated when the next recalculation occurs.
         */
        setDirty(): void;

        /**
         * Displays the card for an active cell if it has rich value content.
         */
        showCard(): void;

        /**
         * Shows the details of the row or column group.
         * @param groupOption Specifies whether to show the details of grouped rows or grouped columns.
         */
        showGroupDetails(groupOption: GroupOption): void;

        /**
         * Ungroups columns and rows for an outline.
         * @param groupOption Specifies how the range can be ungrouped by rows or columns.
         */
        ungroup(groupOption: GroupOption): void;

        /**
         * Unmerge the range cells into separate cells.
         */
        unmerge(): void;

        /**
         * The collection of `ConditionalFormats` that intersect the range.
         */
        getConditionalFormats(): ConditionalFormat[];

        /**
         * Adds a new conditional format to the collection at the first/top priority.
         * @param type The type of conditional format being added. See `ExcelScript.ConditionalFormatType` for details.
         */
        addConditionalFormat(type: ConditionalFormatType): ConditionalFormat;

        /**
         * Clears all conditional formats active on the current specified range.
         */
        clearAllConditionalFormats(): void;

        /**
         * Returns a conditional format for the given ID.
         * @param id The ID of the conditional format.
         */
        getConditionalFormat(id: string): ConditionalFormat;

        /**
         * Represents the cell formula in A1-style notation.
         * If the range contains multiple cells, the data from first cell (represented by row index of 0 and column index of 0) will be returned.
         */
        getFormula(): string;

        /**
         * Sets the cell formula in A1-style notation.
         * If the range contains multiple cells, each cell in the given range will be updated with the input data.
         */
        setFormula(formula: string): void;

        /**
         * Represents the cell formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German.
         * If the range contains multiple cells, the data from first cell (represented by row index of 0 and column index of 0) will be returned.
         */
        getFormulaLocal(): string;

        /**
         * Set the cell formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German.
         * If the range contains multiple cells, each cell in the given range will be updated with the input data.
         */
        setFormulaLocal(formulaLocal: string): void;

        /**
         * Represents the cell formula in R1C1-style notation.
         * If the range contains multiple cells, the data from first cell (represented by row index of 0 and column index of 0) will be returned.
         */
        getFormulaR1C1(): string;

        /**
         * Sets the cell formula in R1C1-style notation.
         * If the range contains multiple cells, each cell in the given range will be updated with the input data.
         */
        setFormulaR1C1(formulaR1C1: string): void;

        /**
         * Represents the data type state of the cell.
         */
        getLinkedDataTypeState(): LinkedDataTypeState;

        /**
         * Represents cell Excel number format code for the given range.
         * If the range contains multiple cells, the data from first cell (represented by row index of 0 and column index of 0) will be returned.
         */
        getNumberFormat(): string;

        /**
         * Sets cell Excel number format code for the given range.
         * If the range contains multiple cells, each cell in the given range will be updated with the input data.
         */
        setNumberFormat(numberFormat: string): void;

        /**
         * Represents cell Excel number format code for the given range, based on the language settings of the user.​
         * Excel does not perform any language or format coercion when getting or setting the `numberFormatLocal` property.
         * Any returned text uses the locally-formatted strings based on the language specified in the system settings.
         * If the range contains multiple cells, the data from first cell (represented by row index of 0 and column index of 0) will be returned.
         */
        getNumberFormatLocal(): string;

        /**
         * Sets cell Excel number format code for the given range, based on the language settings of the user.​
         * Excel does not perform any language or format coercion when getting or setting the `numberFormatLocal` property.
         * Any returned text uses the locally-formatted strings based on the language specified in the system settings.
         * If the range contains multiple cells, each cell in the given range will be updated with the input data.
         */
        setNumberFormatLocal(numberFormatLocal: string): void;

        /**
         * Specifies the number format category of first cell in the range (represented by row index of 0 and column index of 0).
         */
        getNumberFormatCategory(): NumberFormatCategory;

        /**
         * Represents Text value of the specified range. The Text value will not depend on the cell width. The # sign substitution that happens in Excel UI will not affect the text value returned by the API.
         * If the range contains multiple cells, the data from first cell (represented by row index of 0 and column index of 0) will be returned.
         */
        getText(): string;

        /**
         * Represents the type of data in the cell.
         * If the range contains multiple cells, the data from first cell (represented by row index of 0 and column index of 0) will be returned.
         */

        getValueType(): RangeValueType;

        /**
         * Represents the raw value of the specified range. The data returned could be of type string, number, or a boolean. Cell that contain an error will return the error string.
         * If the range contains multiple cells, the data from first cell (represented by row index of 0 and column index of 0) will be returned.
         */
        getValue(): string | number | boolean;

        /**
         * Sets the raw value of the specified range. The data being set could be of type string, number, or a boolean. `null` value will be ignored (not set or overwritten in Excel).
         * If the range contains multiple cells, each cell in the given range will be updated with the input data.
         */
        setValue(value: any): void;
    }

    /**
     * `RangeAreas` represents a collection of one or more rectangular ranges in the same worksheet.
     */
    interface RangeAreas {
        /**
         * Returns the `RangeAreas` reference in A1-style. Address value will contain the worksheet name for each rectangular block of cells (e.g., "Sheet1!A1:B4, Sheet1!D1:D4").
         */
        getAddress(): string;

        /**
         * Returns the `RangeAreas` reference in the user locale.
         */
        getAddressLocal(): string;

        /**
         * Returns the number of rectangular ranges that comprise this `RangeAreas` object.
         */
        getAreaCount(): number;

        /**
         * Returns the number of cells in the `RangeAreas` object, summing up the cell counts of all of the individual rectangular ranges. Returns -1 if the cell count exceeds 2^31-1 (2,147,483,647).
         */
        getCellCount(): number;

        /**
         * Returns a data validation object for all ranges in the `RangeAreas`.
         */
        getDataValidation(): DataValidation;

        /**
         * Returns a `RangeFormat` object, encapsulating the the font, fill, borders, alignment, and other properties for all ranges in the `RangeAreas` object.
         */
        getFormat(): RangeFormat;

        /**
         * Specifies if all the ranges on this `RangeAreas` object represent entire columns (e.g., "A:C, Q:Z").
         */
        getIsEntireColumn(): boolean;

        /**
         * Specifies if all the ranges on this `RangeAreas` object represent entire rows (e.g., "1:3, 5:7").
         */
        getIsEntireRow(): boolean;

        /**
         * Represents the style for all ranges in this `RangeAreas` object.
         * If the styles of the cells are inconsistent, `null` will be returned.
         * For custom styles, the style name will be returned. For built-in styles, a string representing a value in the `BuiltInStyle` enum will be returned.
         */
        getPredefinedCellStyle(): string;

        /**
         * Represents the style for all ranges in this `RangeAreas` object.
         * If the styles of the cells are inconsistent, `null` will be returned.
         * For custom styles, the style name will be returned. For built-in styles, a string representing a value in the `BuiltInStyle` enum will be returned.
         */
        setPredefinedCellStyle(predefinedCellStyle: string): void;

        /**
         * Returns the worksheet for the current `RangeAreas`.
         */
        getWorksheet(): Worksheet;

        /**
         * Calculates all cells in the `RangeAreas`.
         */
        calculate(): void;

        /**
         * Clears values, format, fill, border, and other properties on each of the areas that comprise this `RangeAreas` object.
         * @param applyTo Optional. Determines the type of clear action. See `ExcelScript.ClearApplyTo` for details. Default is "All".
         */
        clear(applyTo?: ClearApplyTo): void;

        /**
         * Converts all cells in the `RangeAreas` with data types into text.
         */
        convertDataTypeToText(): void;

        /**
         * Copies cell data or formatting from the source range or `RangeAreas` to the current `RangeAreas`.
         * The destination `RangeAreas` can be a different size than the source range or `RangeAreas`. The destination will be expanded automatically if it is smaller than the source.
         * @param sourceRange The source range or `RangeAreas` to copy from. When the source `RangeAreas` has multiple ranges, their form must able to be created by removing full rows or columns from a rectangular range.
         * @param copyType The type of cell data or formatting to copy over. Default is "All".
         * @param skipBlanks True if to skip blank cells in the source range or `RangeAreas`. Default is false.
         * @param transpose True if to transpose the cells in the destination `RangeAreas`. Default is false.
         */
        copyFrom(
            sourceRange: Range | RangeAreas | string,
            copyType?: RangeCopyType,
            skipBlanks?: boolean,
            transpose?: boolean
        ): void;

        /**
         * Returns a `RangeAreas` object that represents the entire columns of the `RangeAreas` (for example, if the current `RangeAreas` represents cells "B4:E11, H2", it returns a `RangeAreas` that represents columns "B:E, H:H").
         */
        getEntireColumn(): RangeAreas;

        /**
         * Returns a `RangeAreas` object that represents the entire rows of the `RangeAreas` (for example, if the current `RangeAreas` represents cells "B4:E11", it returns a `RangeAreas` that represents rows "4:11").
         */
        getEntireRow(): RangeAreas;

        /**
         * Returns the `RangeAreas` object that represents the intersection of the given ranges or `RangeAreas`. If no intersection is found, then this method returns `undefined`.
         * @param anotherRange The range, `RangeAreas` object, or address that will be used to determine the intersection.
         */
        getIntersection(anotherRange: Range | RangeAreas | string): RangeAreas;

        /**
         * Returns a `RangeAreas` object that is shifted by the specific row and column offset. The dimension of the returned `RangeAreas` will match the original object. If the resulting `RangeAreas` is forced outside the bounds of the worksheet grid, an error will be thrown.
         * @param rowOffset The number of rows (positive, negative, or 0) by which the `RangeAreas` is to be offset. Positive values are offset downward, and negative values are offset upward.
         * @param columnOffset The number of columns (positive, negative, or 0) by which the `RangeAreas` is to be offset. Positive values are offset to the right, and negative values are offset to the left.
         */
        getOffsetRangeAreas(
            rowOffset: number,
            columnOffset: number
        ): RangeAreas;

        /**
         * Returns a `RangeAreas` object that represents all the cells that match the specified type and value. If no special cells are found that match the criteria, then this method returns `undefined`.
         * @param cellType The type of cells to include.
         * @param cellValueType If `cellType` is either `constants` or `formulas`, this argument is used to determine which types of cells to include in the result. These values can be combined together to return more than one type. The default is to select all constants or formulas, no matter what the type.
         */
        getSpecialCells(
            cellType: SpecialCellType,
            cellValueType?: SpecialCellValueType
        ): RangeAreas;

        /**
         * Returns a scoped collection of tables that overlap with any range in this `RangeAreas` object.
         * @param fullyContained If `true`, returns only tables that are fully contained within the range bounds. Default is `false`.
         */
        getTables(fullyContained?: boolean): Table[];

        /**
         * Returns the used `RangeAreas` that comprises all the used areas of individual rectangular ranges in the `RangeAreas` object.
         * If there are no used cells within the `RangeAreas`, then this method returns `undefined`.
         * @param valuesOnly Whether to only consider cells with values as used cells.
         */
        getUsedRangeAreas(valuesOnly?: boolean): RangeAreas;

        /**
         * Sets the `RangeAreas` to be recalculated when the next recalculation occurs.
         */
        setDirty(): void;

        /**
         * Returns a collection of rectangular ranges that comprise this `RangeAreas` object.
         */
        getAreas(): Range[];

        /**
         * Returns a collection of conditional formats that intersect with any cells in this `RangeAreas` object.
         */
        getConditionalFormats(): ConditionalFormat[];

        /**
         * Adds a new conditional format to the collection at the first/top priority.
         * @param type The type of conditional format being added. See `ExcelScript.ConditionalFormatType` for details.
         */
        addConditionalFormat(type: ConditionalFormatType): ConditionalFormat;

        /**
         * Clears all conditional formats active on the current specified range.
         */
        clearAllConditionalFormats(): void;

        /**
         * Returns a conditional format for the given ID.
         * @param id The ID of the conditional format.
         */
        getConditionalFormat(id: string): ConditionalFormat;
    }

    /**
     * Represents a collection of one or more rectangular ranges in multiple worksheets.
     */
    interface WorkbookRangeAreas {
        /**
         * Returns an array of addresses in A1-style. Address values contain the worksheet name for each rectangular block of cells (e.g., "Sheet1!A1:B4, Sheet1!D1:D4"). Read-only.
         */
        getAddresses(): string[];

        /**
         * Returns the `RangeAreas` object based on worksheet name or ID in the collection. If the worksheet does not exist, then this method returns `undefined`.
         * @param key The name or ID of the worksheet.
         */
        getRangeAreasBySheet(key: string): RangeAreas;

        /**
         * Returns the `RangeAreasCollection` object. Each `RangeAreas` in the collection represent one or more rectangle ranges in one worksheet.
         */
        getAreas(): RangeAreas[];

        /**
         * Returns ranges that comprise this object in a `RangeCollection` object.
         */
        getRanges(): Range[];
    }

    /**
     * RangeView represents a set of visible cells of the parent range.
     */
    interface RangeView {
        /**
         * Represents the cell addresses of the `RangeView`.
         */
        getCellAddresses(): string[][];

        /**
         * The number of visible columns.
         */
        getColumnCount(): number;

        /**
         * Represents the formula in A1-style notation. If a cell has no formula, its value is returned instead.
         */
        getFormulas(): string[][];

        /**
         * Represents the formula in A1-style notation. If a cell has no formula, its value is returned instead.
         */
        setFormulas(formulas: string[][]): void;

        /**
         * Represents the formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German. If a cell has no formula, its value is returned instead.
         */
        getFormulasLocal(): string[][];

        /**
         * Represents the formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German. If a cell has no formula, its value is returned instead.
         */
        setFormulasLocal(formulasLocal: string[][]): void;

        /**
         * Represents the formula in R1C1-style notation. If a cell has no formula, its value is returned instead.
         */
        getFormulasR1C1(): string[][];

        /**
         * Represents the formula in R1C1-style notation. If a cell has no formula, its value is returned instead.
         */
        setFormulasR1C1(formulasR1C1: string[][]): void;

        /**
         * Returns a value that represents the index of the `RangeView`.
         */
        getIndex(): number;

        /**
         * Represents Excel's number format code for the given cell.
         */
        getNumberFormat(): string[][];

        /**
         * Represents Excel's number format code for the given cell.
         */
        setNumberFormat(numberFormat: string[][]): void;

        /**
         * The number of visible rows.
         */
        getRowCount(): number;

        /**
         * Text values of the specified range. The text value will not depend on the cell width. The # sign substitution that happens in Excel UI will not affect the text value returned by the API.
         */
        getText(): string[][];

        /**
         * Represents the type of data of each cell.
         */
        getValueTypes(): RangeValueType[][];

        /**
         * Represents the raw values of the specified range view. The data returned could be of type string, number, or a boolean. Cells that contain an error will return the error string.
         */
        getValues(): (string | number | boolean)[][];

        /**
         * Represents the raw values of the specified range view. The data returned could be of type string, number, or a boolean. Cells that contain an error will return the error string.
         */
        setValues(values: (string | number | boolean)[][]): void;

        /**
         * Gets the parent range associated with the current `RangeView`.
         */
        getRange(): Range;

        /**
         * Represents a collection of range views associated with the range.
         */
        getRows(): RangeView[];
    }

    /**
     * Represents a defined name for a range of cells or value. Names can be primitive named objects (as seen in the type below), range object, or a reference to a range. This object can be used to obtain range object associated with names.
     */
    interface NamedItem {
        /**
         * Returns an object containing values and types of the named item.
         */
        getArrayValues(): NamedItemArrayValues;

        /**
         * Specifies the comment associated with this name.
         */
        getComment(): string;

        /**
         * Specifies the comment associated with this name.
         */
        setComment(comment: string): void;

        /**
         * The formula of the named item. Formulas always start with an equal sign ("=").
         */
        getFormula(): string;

        /**
         * The formula of the named item. Formulas always start with an equal sign ("=").
         */
        setFormula(formula: string): void;

        /**
         * The name of the object.
         */
        getName(): string;

        /**
         * Specifies if the name is scoped to the workbook or to a specific worksheet. Possible values are: Worksheet, Workbook.
         */
        getScope(): NamedItemScope;

        /**
         * Specifies the type of the value returned by the name's formula. See `ExcelScript.NamedItemType` for details.
         */
        getType(): NamedItemType;

        /**
         * Represents the value computed by the name's formula. For a named range, will return the range address.
         */
        getValue(): string | number;

        /**
         * Specifies if the object is visible.
         */
        getVisible(): boolean;

        /**
         * Specifies if the object is visible.
         */
        setVisible(visible: boolean): void;

        /**
         * Returns the worksheet to which the named item is scoped. If the item is scoped to the workbook instead, then this method returns `undefined`.
         */
        getWorksheet(): Worksheet | undefined;

        /**
         * Deletes the given name.
         */
        delete(): void;

        /**
         * Returns the range object that is associated with the name. If the named item's type is not a range, then this method returns `undefined`.
         */
        getRange(): Range;
    }

    /**
     * Represents an object containing values and types of a named item.
     */
    interface NamedItemArrayValues {
        /**
         * Represents the types for each item in the named item array
         */
        getTypes(): RangeValueType[][];

        /**
         * Represents the values of each item in the named item array.
         */
        getValues(): (string | number | boolean)[][];
    }

    /**
     * Represents an Office.js binding that is defined in the workbook.
     */
    interface Binding {
        /**
         * Represents the binding identifier.
         */
        getId(): string;

        /**
         * Returns the type of the binding. See `ExcelScript.BindingType` for details.
         */
        getType(): BindingType;

        /**
         * Deletes the binding.
         */
        delete(): void;

        /**
         * Returns the range represented by the binding. Will throw an error if the binding is not of the correct type.
         */
        getRange(): Range;

        /**
         * Returns the table represented by the binding. Will throw an error if the binding is not of the correct type.
         */
        getTable(): Table;

        /**
         * Returns the text represented by the binding. Will throw an error if the binding is not of the correct type.
         */
        getText(): string;
    }

    /**
     * Represents an Excel table.
     */
    interface Table {
        /**
         * Represents the `AutoFilter` object of the table.
         */
        getAutoFilter(): AutoFilter;

        /**
         * Specifies if the first column contains special formatting.
         */
        getHighlightFirstColumn(): boolean;

        /**
         * Specifies if the first column contains special formatting.
         */
        setHighlightFirstColumn(highlightFirstColumn: boolean): void;

        /**
         * Specifies if the last column contains special formatting.
         */
        getHighlightLastColumn(): boolean;

        /**
         * Specifies if the last column contains special formatting.
         */
        setHighlightLastColumn(highlightLastColumn: boolean): void;

        /**
         * Returns a value that uniquely identifies the table in a given workbook. The value of the identifier remains the same even when the table is renamed.
         */
        getId(): string;

        /**
         * Returns a numeric ID.
         */
        getLegacyId(): string;

        /**
         * Name of the table.
         *
         */
        getName(): string;

        /**
         * Name of the table.
         *
         */
        setName(name: string): void;

        /**
         * Specifies if the columns show banded formatting in which odd columns are highlighted differently from even ones, to make reading the table easier.
         */
        getShowBandedColumns(): boolean;

        /**
         * Specifies if the columns show banded formatting in which odd columns are highlighted differently from even ones, to make reading the table easier.
         */
        setShowBandedColumns(showBandedColumns: boolean): void;

        /**
         * Specifies if the rows show banded formatting in which odd rows are highlighted differently from even ones, to make reading the table easier.
         */
        getShowBandedRows(): boolean;

        /**
         * Specifies if the rows show banded formatting in which odd rows are highlighted differently from even ones, to make reading the table easier.
         */
        setShowBandedRows(showBandedRows: boolean): void;

        /**
         * Specifies if the filter buttons are visible at the top of each column header. Setting this is only allowed if the table contains a header row.
         */
        getShowFilterButton(): boolean;

        /**
         * Specifies if the filter buttons are visible at the top of each column header. Setting this is only allowed if the table contains a header row.
         */
        setShowFilterButton(showFilterButton: boolean): void;

        /**
         * Specifies if the header row is visible. This value can be set to show or remove the header row.
         */
        getShowHeaders(): boolean;

        /**
         * Specifies if the header row is visible. This value can be set to show or remove the header row.
         */
        setShowHeaders(showHeaders: boolean): void;

        /**
         * Specifies if the total row is visible. This value can be set to show or remove the total row.
         */
        getShowTotals(): boolean;

        /**
         * Specifies if the total row is visible. This value can be set to show or remove the total row.
         */
        setShowTotals(showTotals: boolean): void;

        /**
         * Represents the sorting for the table.
         */
        getSort(): TableSort;

        /**
         * Constant value that represents the table style. Possible values are: "TableStyleLight1" through "TableStyleLight21", "TableStyleMedium1" through "TableStyleMedium28", "TableStyleDark1" through "TableStyleDark11". A custom user-defined style present in the workbook can also be specified.
         */
        getPredefinedTableStyle(): string;

        /**
         * Constant value that represents the table style. Possible values are: "TableStyleLight1" through "TableStyleLight21", "TableStyleMedium1" through "TableStyleMedium28", "TableStyleDark1" through "TableStyleDark11". A custom user-defined style present in the workbook can also be specified.
         */
        setPredefinedTableStyle(predefinedTableStyle: string): void;

        /**
         * The worksheet containing the current table.
         */
        getWorksheet(): Worksheet;

        /**
         * Clears all the filters currently applied on the table.
         */
        clearFilters(): void;

        /**
         * Converts the table into a normal range of cells. All data is preserved.
         */
        convertToRange(): Range;

        /**
         * Deletes the table.
         */
        delete(): void;

        /**
         * Gets the range object associated with the data body of the table.
         */
        getRangeBetweenHeaderAndTotal(): Range;

        /**
         * Gets the range object associated with the header row of the table.
         */
        getHeaderRowRange(): Range;

        /**
         * Gets the range object associated with the entire table.
         */
        getRange(): Range;

        /**
         * Gets the range object associated with the totals row of the table.
         */
        getTotalRowRange(): Range;

        /**
         * Reapplies all the filters currently on the table.
         */
        reapplyFilters(): void;

        /**
         * Resize the table to the new range. The new range must overlap with the original table range and the headers (or the top of the table) must be in the same row.
         * @param newRange The range object or range address that will be used to determine the new size of the table.
         */
        resize(newRange: Range | string): void;

        /**
         * Represents a collection of all the columns in the table.
         */
        getColumns(): TableColumn[];

        /**
         * Gets a column object by name or ID. If the column doesn't exist, then this method returns `undefined`.
         * @param key Column name or ID.
         */
        getColumn(key: number | string): TableColumn | undefined;

        /**
         * Adds one row to the table.
         * @param index Optional. Specifies the relative position of the new row. If null or -1, the addition happens at the end. Any rows below the inserted row are shifted downwards. Zero-indexed.
         * @param values Optional. A 1-dimensional array of unformatted values of the table row.
         */
        addRow(index?: number, values?: (boolean | string | number)[]): void;

        /**
         * Adds one or more rows to the table.
         * @param index Optional. Specifies the relative position of the new row. If null or -1, the addition happens at the end. Any rows below the inserted row are shifted downwards. Zero-indexed.
         * @param values Optional. A 2-dimensional array of unformatted values of the table row.
         */
        addRows(index?: number, values?: (boolean | string | number)[][]): void;

        /**
         * Adds a new column to the table.
         * @param index Optional. Specifies the relative position of the new column. If null or -1, the addition happens at the end. Columns with a higher index will be shifted to the side. Zero-indexed.
         * @param values Optional. A 1-dimensional array of unformatted values of the table column.
         * @param name Optional. Specifies the name of the new column. If null, the default name will be used.
         */
        addColumn(
            index?: number,
            values?: (boolean | string | number)[],
            name?: string
        ): TableColumn;

        /**
         * Delete a specified number of rows at a given index.
         * @param index The index value of the row to be deleted. Caution: the index of the row may have moved from the time you determined the value to use for removal.
         * @param count Number of rows to delete. By default, a single row will be deleted. Note: Deleting more than 1000 rows at the same time could result in a Power Automate timeout.
         */
        deleteRowsAt(index: number, count?: number): void;

        /**
         * Gets a column object by ID. If the column does not exist, will return undefined.
         * @param key Column ID.
         */
        getColumnById(key: number): TableColumn | undefined;

        /**
         * Gets a column object by Name. If the column does not exist, will return undefined.
         * @param key Column Name.
         */
        getColumnByName(key: string): TableColumn | undefined;

        /**
         * Gets the number of rows in the table.
         */
        getRowCount(): number;
    }

    /**
     * Represents a column in a table.
     */
    interface TableColumn {
        /**
         * Retrieves the filter applied to the column.
         */
        getFilter(): Filter;

        /**
         * Returns a unique key that identifies the column within the table.
         */
        getId(): number;

        /**
         * Returns the index number of the column within the columns collection of the table. Zero-indexed.
         */
        getIndex(): number;

        /**
         * Specifies the name of the table column.
         */
        getName(): string;

        /**
         * Specifies the name of the table column.
         */
        setName(name: string): void;

        /**
         * Deletes the column from the table.
         */
        delete(): void;

        /**
         * Gets the range object associated with the data body of the column.
         */
        getRangeBetweenHeaderAndTotal(): Range;

        /**
         * Gets the range object associated with the header row of the column.
         */
        getHeaderRowRange(): Range;

        /**
         * Gets the range object associated with the entire column.
         */
        getRange(): Range;

        /**
         * Gets the range object associated with the totals row of the column.
         */
        getTotalRowRange(): Range;
    }

    /**
     * Represents the data validation applied to the current range.
     */
    interface DataValidation {
        /**
         * Error alert when user enters invalid data.
         */
        getErrorAlert(): DataValidationErrorAlert;

        /**
         * Error alert when user enters invalid data.
         */
        setErrorAlert(errorAlert: DataValidationErrorAlert): void;

        /**
         * Specifies if data validation will be performed on blank cells. Default is `true`.
         */
        getIgnoreBlanks(): boolean;

        /**
         * Specifies if data validation will be performed on blank cells. Default is `true`.
         */
        setIgnoreBlanks(ignoreBlanks: boolean): void;

        /**
         * Prompt when users select a cell.
         */
        getPrompt(): DataValidationPrompt;

        /**
         * Prompt when users select a cell.
         */
        setPrompt(prompt: DataValidationPrompt): void;

        /**
         * Data validation rule that contains different type of data validation criteria.
         */
        getRule(): DataValidationRule;

        /**
         * Data validation rule that contains different type of data validation criteria.
         */
        setRule(rule: DataValidationRule): void;

        /**
         * Type of the data validation, see `ExcelScript.DataValidationType` for details.
         */
        getType(): DataValidationType;

        /**
         * Represents if all cell values are valid according to the data validation rules.
         * Returns `true` if all cell values are valid, or `false` if all cell values are invalid.
         * Returns `null` if there are both valid and invalid cell values within the range.
         */
        getValid(): boolean;

        /**
         * Clears the data validation from the current range.
         */
        clear(): void;

        /**
         * Returns a `RangeAreas` object, comprising one or more rectangular ranges, with invalid cell values. If all cell values are valid, this method will return `null`.
         */
        getInvalidCells(): RangeAreas;
    }

    /**
     * Represents the results from `Range.removeDuplicates`.
     */
    interface RemoveDuplicatesResult {
        /**
         * Number of duplicated rows removed by the operation.
         */
        getRemoved(): number;

        /**
         * Number of remaining unique rows present in the resulting range.
         */
        getUniqueRemaining(): number;
    }

    /**
     * A format object encapsulating the range's font, fill, borders, alignment, and other properties.
     */
    interface RangeFormat {
        /**
         * Specifies if text is automatically indented when text alignment is set to equal distribution.
         */
        getAutoIndent(): boolean;

        /**
         * Specifies if text is automatically indented when text alignment is set to equal distribution.
         */
        setAutoIndent(autoIndent: boolean): void;

        /**
         * Specifies the width of all colums within the range. If the column widths are not uniform, `null` will be returned.
         */
        getColumnWidth(): number;

        /**
         * Specifies the width of all colums within the range. If the column widths are not uniform, `null` will be returned.
         */
        setColumnWidth(columnWidth: number): void;

        /**
         * Returns the fill object defined on the overall range.
         */
        getFill(): RangeFill;

        /**
         * Returns the font object defined on the overall range.
         */
        getFont(): RangeFont;

        /**
         * Represents the horizontal alignment for the specified object. See `ExcelScript.HorizontalAlignment` for details.
         */
        getHorizontalAlignment(): HorizontalAlignment;

        /**
         * Represents the horizontal alignment for the specified object. See `ExcelScript.HorizontalAlignment` for details.
         */
        setHorizontalAlignment(horizontalAlignment: HorizontalAlignment): void;

        /**
         * An integer from 0 to 250 that indicates the indent level.
         */
        getIndentLevel(): number;

        /**
         * An integer from 0 to 250 that indicates the indent level.
         */
        setIndentLevel(indentLevel: number): void;

        /**
         * Returns the format protection object for a range.
         */
        getProtection(): FormatProtection;

        /**
         * The reading order for the range.
         */
        getReadingOrder(): ReadingOrder;

        /**
         * The reading order for the range.
         */
        setReadingOrder(readingOrder: ReadingOrder): void;

        /**
         * The height of all rows in the range. If the row heights are not uniform, `null` will be returned.
         */
        getRowHeight(): number;

        /**
         * The height of all rows in the range. If the row heights are not uniform, `null` will be returned.
         */
        setRowHeight(rowHeight: number): void;

        /**
         * Specifies if text automatically shrinks to fit in the available column width.
         */
        getShrinkToFit(): boolean;

        /**
         * Specifies if text automatically shrinks to fit in the available column width.
         */
        setShrinkToFit(shrinkToFit: boolean): void;

        /**
         * The text orientation of all the cells within the range.
         * The text orientation should be an integer either from -90 to 90, or 180 for vertically-oriented text.
         * If the orientation within a range are not uniform, then `null` will be returned.
         */
        getTextOrientation(): number;

        /**
         * The text orientation of all the cells within the range.
         * The text orientation should be an integer either from -90 to 90, or 180 for vertically-oriented text.
         * If the orientation within a range are not uniform, then `null` will be returned.
         */
        setTextOrientation(textOrientation: number): void;

        /**
         * Determines if the row height of the `Range` object equals the standard height of the sheet.
         * Returns `true` if the row height of the `Range` object equals the standard height of the sheet.
         * Returns `null` if the range contains more than one row and the rows aren't all the same height.
         * Returns `false` otherwise.
         * Note: This property is only intended to be set to `true`. Setting it to `false` has no effect.
         */
        getUseStandardHeight(): boolean;

        /**
         * Determines if the row height of the `Range` object equals the standard height of the sheet.
         * Returns `true` if the row height of the `Range` object equals the standard height of the sheet.
         * Returns `null` if the range contains more than one row and the rows aren't all the same height.
         * Returns `false` otherwise.
         * Note: This property is only intended to be set to `true`. Setting it to `false` has no effect.
         */
        setUseStandardHeight(useStandardHeight: boolean): void;

        /**
         * Specifies if the column width of the `Range` object equals the standard width of the sheet.
         * Returns `true` if the column width of the `Range` object equals the standard width of the sheet.
         * Returns `null` if the range contains more than one column and the columns aren't all the same height.
         * Returns `false` otherwise.
         * Note: This property is only intended to be set to `true`. Setting it to `false` has no effect.
         */
        getUseStandardWidth(): boolean;

        /**
         * Specifies if the column width of the `Range` object equals the standard width of the sheet.
         * Returns `true` if the column width of the `Range` object equals the standard width of the sheet.
         * Returns `null` if the range contains more than one column and the columns aren't all the same height.
         * Returns `false` otherwise.
         * Note: This property is only intended to be set to `true`. Setting it to `false` has no effect.
         */
        setUseStandardWidth(useStandardWidth: boolean): void;

        /**
         * Represents the vertical alignment for the specified object. See `ExcelScript.VerticalAlignment` for details.
         */
        getVerticalAlignment(): VerticalAlignment;

        /**
         * Represents the vertical alignment for the specified object. See `ExcelScript.VerticalAlignment` for details.
         */
        setVerticalAlignment(verticalAlignment: VerticalAlignment): void;

        /**
         * Specifies if Excel wraps the text in the object. A `null` value indicates that the entire range doesn't have a uniform wrap setting
         */
        getWrapText(): boolean;

        /**
         * Specifies if Excel wraps the text in the object. A `null` value indicates that the entire range doesn't have a uniform wrap setting
         */
        setWrapText(wrapText: boolean): void;

        /**
         * Adjusts the indentation of the range formatting. The indent value ranges from 0 to 250 and is measured in characters.
         * @param amount The number of character spaces by which the current indent is adjusted. This value should be between -250 and 250.
         * **Note**: If the amount would raise the indent level above 250, the indent level stays with 250.
         * Similarly, if the amount would lower the indent level below 0, the indent level stays 0.
         */
        adjustIndent(amount: number): void;

        /**
         * Changes the width of the columns of the current range to achieve the best fit, based on the current data in the columns.
         */
        autofitColumns(): void;

        /**
         * Changes the height of the rows of the current range to achieve the best fit, based on the current data in the columns.
         */
        autofitRows(): void;

        /**
         * Collection of border objects that apply to the overall range.
         */
        getBorders(): RangeBorder[];

        /**
         * Specifies a double that lightens or darkens a color for range borders. The value is between -1 (darkest) and 1 (brightest), with 0 for the original color.
         * A `null` value indicates that the entire border collection doesn't have a uniform `tintAndShade` setting.
         */
        getRangeBorderTintAndShade(): number;

        /**
         * Specifies a double that lightens or darkens a color for range borders. The value is between -1 (darkest) and 1 (brightest), with 0 for the original color.
         * A `null` value indicates that the entire border collection doesn't have a uniform `tintAndShade` setting.
         */
        setRangeBorderTintAndShade(rangeBorderTintAndShade: number): void;

        /**
         * Gets a border object using its name.
         * @param index Index value of the border object to be retrieved. See `ExcelScript.BorderIndex` for details.
         */
        getRangeBorder(index: BorderIndex): RangeBorder;
    }

    /**
     * Represents the format protection of a range object.
     */
    interface FormatProtection {
        /**
         * Specifies if Excel hides the formula for the cells in the range. A `null` value indicates that the entire range doesn't have a uniform formula hidden setting.
         */
        getFormulaHidden(): boolean;

        /**
         * Specifies if Excel hides the formula for the cells in the range. A `null` value indicates that the entire range doesn't have a uniform formula hidden setting.
         */
        setFormulaHidden(formulaHidden: boolean): void;

        /**
         * Specifies if Excel locks the cells in the object. A `null` value indicates that the entire range doesn't have a uniform lock setting.
         */
        getLocked(): boolean;

        /**
         * Specifies if Excel locks the cells in the object. A `null` value indicates that the entire range doesn't have a uniform lock setting.
         */
        setLocked(locked: boolean): void;
    }

    /**
     * Represents the background of a range object.
     */
    interface RangeFill {
        /**
         * HTML color code representing the color of the background, in the form #RRGGBB (e.g., "FFA500"), or as a named HTML color (e.g., "orange")
         */
        getColor(): string;

        /**
         * HTML color code representing the color of the background, in the form #RRGGBB (e.g., "FFA500"), or as a named HTML color (e.g., "orange")
         */
        setColor(color: string): void;

        /**
         * The pattern of a range. See `ExcelScript.FillPattern` for details. LinearGradient and RectangularGradient are not supported.
         * A `null` value indicates that the entire range doesn't have a uniform pattern setting.
         */
        getPattern(): FillPattern;

        /**
         * The pattern of a range. See `ExcelScript.FillPattern` for details. LinearGradient and RectangularGradient are not supported.
         * A `null` value indicates that the entire range doesn't have a uniform pattern setting.
         */
        setPattern(pattern: FillPattern): void;

        /**
         * The HTML color code representing the color of the range pattern, in the form #RRGGBB (e.g., "FFA500"), or as a named HTML color (e.g., "orange").
         */
        getPatternColor(): string;

        /**
         * The HTML color code representing the color of the range pattern, in the form #RRGGBB (e.g., "FFA500"), or as a named HTML color (e.g., "orange").
         */
        setPatternColor(patternColor: string): void;

        /**
         * Specifies a double that lightens or darkens a pattern color for the range fill. The value is between -1 (darkest) and 1 (brightest), with 0 for the original color.
         * A `null` value indicates that the range doesn't have uniform `patternTintAndShade` settings.
         */
        getPatternTintAndShade(): number;

        /**
         * Specifies a double that lightens or darkens a pattern color for the range fill. The value is between -1 (darkest) and 1 (brightest), with 0 for the original color.
         * A `null` value indicates that the range doesn't have uniform `patternTintAndShade` settings.
         */
        setPatternTintAndShade(patternTintAndShade: number): void;

        /**
         * Specifies a double that lightens or darkens a color for the range fill. The value is between -1 (darkest) and 1 (brightest), with 0 for the original color.
         * A `null` value indicates that the range doesn't have uniform `tintAndShade` settings.
         */
        getTintAndShade(): number;

        /**
         * Specifies a double that lightens or darkens a color for the range fill. The value is between -1 (darkest) and 1 (brightest), with 0 for the original color.
         * A `null` value indicates that the range doesn't have uniform `tintAndShade` settings.
         */
        setTintAndShade(tintAndShade: number): void;

        /**
         * Resets the range background.
         */
        clear(): void;
    }

    /**
     * Represents the border of an object.
     */
    interface RangeBorder {
        /**
         * HTML color code representing the color of the border line, in the form #RRGGBB (e.g., "FFA500"), or as a named HTML color (e.g., "orange").
         */
        getColor(): string;

        /**
         * HTML color code representing the color of the border line, in the form #RRGGBB (e.g., "FFA500"), or as a named HTML color (e.g., "orange").
         */
        setColor(color: string): void;

        /**
         * Constant value that indicates the specific side of the border. See `ExcelScript.BorderIndex` for details.
         */
        getSideIndex(): BorderIndex;

        /**
         * One of the constants of line style specifying the line style for the border. See `ExcelScript.BorderLineStyle` for details.
         */
        getStyle(): BorderLineStyle;

        /**
         * One of the constants of line style specifying the line style for the border. See `ExcelScript.BorderLineStyle` for details.
         */
        setStyle(style: BorderLineStyle): void;

        /**
         * Specifies a double that lightens or darkens a color for the range border, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.
         * A `null` value indicates that the border doesn't have a uniform `tintAndShade` setting.
         */
        getTintAndShade(): number;

        /**
         * Specifies a double that lightens or darkens a color for the range border, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.
         * A `null` value indicates that the border doesn't have a uniform `tintAndShade` setting.
         */
        setTintAndShade(tintAndShade: number): void;

        /**
         * Specifies the weight of the border around a range. See `ExcelScript.BorderWeight` for details.
         */
        getWeight(): BorderWeight;

        /**
         * Specifies the weight of the border around a range. See `ExcelScript.BorderWeight` for details.
         */
        setWeight(weight: BorderWeight): void;
    }

    /**
     * This object represents the font attributes (font name, font size, color, etc.) for an object.
     */
    interface RangeFont {
        /**
         * Represents the bold status of the font.
         */
        getBold(): boolean;

        /**
         * Represents the bold status of the font.
         */
        setBold(bold: boolean): void;

        /**
         * HTML color code representation of the text color (e.g., #FF0000 represents Red).
         */
        getColor(): string;

        /**
         * HTML color code representation of the text color (e.g., #FF0000 represents Red).
         */
        setColor(color: string): void;

        /**
         * Specifies the italic status of the font.
         */
        getItalic(): boolean;

        /**
         * Specifies the italic status of the font.
         */
        setItalic(italic: boolean): void;

        /**
         * Font name (e.g., "Calibri"). The name's length should not be greater than 31 characters.
         */
        getName(): string;

        /**
         * Font name (e.g., "Calibri"). The name's length should not be greater than 31 characters.
         */
        setName(name: string): void;

        /**
         * Font size.
         */
        getSize(): number;

        /**
         * Font size.
         */
        setSize(size: number): void;

        /**
         * Specifies the strikethrough status of font. A `null` value indicates that the entire range doesn't have a uniform strikethrough setting.
         */
        getStrikethrough(): boolean;

        /**
         * Specifies the strikethrough status of font. A `null` value indicates that the entire range doesn't have a uniform strikethrough setting.
         */
        setStrikethrough(strikethrough: boolean): void;

        /**
         * Specifies the subscript status of font.
         * Returns `true` if all the fonts of the range are subscript.
         * Returns `false` if all the fonts of the range are superscript or normal (neither superscript, nor subscript).
         * Returns `null` otherwise.
         */
        getSubscript(): boolean;

        /**
         * Specifies the subscript status of font.
         * Returns `true` if all the fonts of the range are subscript.
         * Returns `false` if all the fonts of the range are superscript or normal (neither superscript, nor subscript).
         * Returns `null` otherwise.
         */
        setSubscript(subscript: boolean): void;

        /**
         * Specifies the superscript status of font.
         * Returns `true` if all the fonts of the range are superscript.
         * Returns `false` if all the fonts of the range are subscript or normal (neither superscript, nor subscript).
         * Returns `null` otherwise.
         */
        getSuperscript(): boolean;

        /**
         * Specifies the superscript status of font.
         * Returns `true` if all the fonts of the range are superscript.
         * Returns `false` if all the fonts of the range are subscript or normal (neither superscript, nor subscript).
         * Returns `null` otherwise.
         */
        setSuperscript(superscript: boolean): void;

        /**
         * Specifies a double that lightens or darkens a color for the range font. The value is between -1 (darkest) and 1 (brightest), with 0 for the original color.
         * A `null` value indicates that the entire range doesn't have a uniform font `tintAndShade` setting.
         */
        getTintAndShade(): number;

        /**
         * Specifies a double that lightens or darkens a color for the range font. The value is between -1 (darkest) and 1 (brightest), with 0 for the original color.
         * A `null` value indicates that the entire range doesn't have a uniform font `tintAndShade` setting.
         */
        setTintAndShade(tintAndShade: number): void;

        /**
         * Type of underline applied to the font. See `ExcelScript.RangeUnderlineStyle` for details.
         */
        getUnderline(): RangeUnderlineStyle;

        /**
         * Type of underline applied to the font. See `ExcelScript.RangeUnderlineStyle` for details.
         */
        setUnderline(underline: RangeUnderlineStyle): void;
    }

    /**
     * Represents a chart object in a workbook.
     */
    interface Chart {
        /**
         * Represents chart axes.
         */
        getAxes(): ChartAxes;

        /**
         * Specifies a chart category label level enumeration constant, referring to the level of the source category labels.
         */
        getCategoryLabelLevel(): number;

        /**
         * Specifies a chart category label level enumeration constant, referring to the level of the source category labels.
         */
        setCategoryLabelLevel(categoryLabelLevel: number): void;

        /**
         * Specifies the type of the chart. See `ExcelScript.ChartType` for details.
         */
        getChartType(): ChartType;

        /**
         * Specifies the type of the chart. See `ExcelScript.ChartType` for details.
         */
        setChartType(chartType: ChartType): void;

        /**
         * Represents the data labels on the chart.
         */
        getDataLabels(): ChartDataLabels;

        /**
         * Specifies the way that blank cells are plotted on a chart.
         */
        getDisplayBlanksAs(): ChartDisplayBlanksAs;

        /**
         * Specifies the way that blank cells are plotted on a chart.
         */
        setDisplayBlanksAs(displayBlanksAs: ChartDisplayBlanksAs): void;

        /**
         * Encapsulates the format properties for the chart area.
         */
        getFormat(): ChartAreaFormat;

        /**
         * Specifies the height, in points, of the chart object.
         */
        getHeight(): number;

        /**
         * Specifies the height, in points, of the chart object.
         */
        setHeight(height: number): void;

        /**
         * The unique ID of chart.
         */
        getId(): string;

        /**
         * The distance, in points, from the left side of the chart to the worksheet origin.
         */
        getLeft(): number;

        /**
         * The distance, in points, from the left side of the chart to the worksheet origin.
         */
        setLeft(left: number): void;

        /**
         * Represents the legend for the chart.
         */
        getLegend(): ChartLegend;

        /**
         * Specifies the name of a chart object.
         */
        getName(): string;

        /**
         * Specifies the name of a chart object.
         */
        setName(name: string): void;

        /**
         * Encapsulates the options for a pivot chart.
         */
        getPivotOptions(): ChartPivotOptions;

        /**
         * Represents the plot area for the chart.
         */
        getPlotArea(): ChartPlotArea;

        /**
         * Specifies the way columns or rows are used as data series on the chart.
         */
        getPlotBy(): ChartPlotBy;

        /**
         * Specifies the way columns or rows are used as data series on the chart.
         */
        setPlotBy(plotBy: ChartPlotBy): void;

        /**
         * True if only visible cells are plotted. False if both visible and hidden cells are plotted.
         */
        getPlotVisibleOnly(): boolean;

        /**
         * True if only visible cells are plotted. False if both visible and hidden cells are plotted.
         */
        setPlotVisibleOnly(plotVisibleOnly: boolean): void;

        /**
         * Specifies a chart series name level enumeration constant, referring to the level of the source series names.
         */
        getSeriesNameLevel(): number;

        /**
         * Specifies a chart series name level enumeration constant, referring to the level of the source series names.
         */
        setSeriesNameLevel(seriesNameLevel: number): void;

        /**
         * Specifies whether to display all field buttons on a PivotChart.
         */
        getShowAllFieldButtons(): boolean;

        /**
         * Specifies whether to display all field buttons on a PivotChart.
         */
        setShowAllFieldButtons(showAllFieldButtons: boolean): void;

        /**
         * Specifies whether to show the data labels when the value is greater than the maximum value on the value axis.
         * If the value axis becomes smaller than the size of the data points, you can use this property to set whether to show the data labels.
         * This property applies to 2-D charts only.
         */
        getShowDataLabelsOverMaximum(): boolean;

        /**
         * Specifies whether to show the data labels when the value is greater than the maximum value on the value axis.
         * If the value axis becomes smaller than the size of the data points, you can use this property to set whether to show the data labels.
         * This property applies to 2-D charts only.
         */
        setShowDataLabelsOverMaximum(showDataLabelsOverMaximum: boolean): void;

        /**
         * Specifies the chart style for the chart.
         */
        getStyle(): number;

        /**
         * Specifies the chart style for the chart.
         */
        setStyle(style: number): void;

        /**
         * Represents the title of the specified chart, including the text, visibility, position, and formatting of the title.
         */
        getTitle(): ChartTitle;

        /**
         * Specifies the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).
         */
        getTop(): number;

        /**
         * Specifies the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).
         */
        setTop(top: number): void;

        /**
         * Specifies the width, in points, of the chart object.
         */
        getWidth(): number;

        /**
         * Specifies the width, in points, of the chart object.
         */
        setWidth(width: number): void;

        /**
         * The worksheet containing the current chart.
         */
        getWorksheet(): Worksheet;

        /**
         * Activates the chart in the Excel UI.
         */
        activate(): void;

        /**
         * Deletes the chart object.
         */
        delete(): void;

        /**
         * Renders the chart as a base64-encoded image by scaling the chart to fit the specified dimensions.
         * The aspect ratio is preserved as part of the resizing.
         * @param height Optional. The desired height of the resulting image.
         * @param width Optional. The desired width of the resulting image.
         * @param fittingMode Optional. The method used to scale the chart to the specified dimensions (if both height and width are set).
         */
        getImage(
            width?: number,
            height?: number,
            fittingMode?: ImageFittingMode
        ): string;

        /**
         * Resets the source data for the chart.
         * @param sourceData The range object corresponding to the source data.
         * @param seriesBy Specifies the way columns or rows are used as data series on the chart. Can be one of the following: Auto (default), Rows, and Columns. See `ExcelScript.ChartSeriesBy` for details.
         */
        setData(sourceData: Range, seriesBy?: ChartSeriesBy): void;

        /**
         * Positions the chart relative to cells on the worksheet.
         * @param startCell The start cell. This is where the chart will be moved to. The start cell is the top-left or top-right cell, depending on the user's right-to-left display settings.
         * @param endCell Optional. The end cell. If specified, the chart's width and height will be set to fully cover up this cell/range.
         */
        setPosition(startCell: Range | string, endCell?: Range | string): void;

        /**
         * Represents either a single series or collection of series in the chart.
         */
        getSeries(): ChartSeries[];

        /**
         * Add a new series to the collection. The new added series is not visible until values, x-axis values, or bubble sizes for it are set (depending on chart type).
         * @param name Optional. Name of the series.
         * @param index Optional. Index value of the series to be added. Zero-indexed.
         */
        addChartSeries(name?: string, index?: number): ChartSeries;
    }

    /**
     * Encapsulates the options for the pivot chart.
     */
    interface ChartPivotOptions {
        /**
         * Specifies whether to display the axis field buttons on a PivotChart. The `showAxisFieldButtons` property corresponds to the "Show Axis Field Buttons" command on the "Field Buttons" drop-down list of the "Analyze" tab, which is available when a PivotChart is selected.
         */
        getShowAxisFieldButtons(): boolean;

        /**
         * Specifies whether to display the axis field buttons on a PivotChart. The `showAxisFieldButtons` property corresponds to the "Show Axis Field Buttons" command on the "Field Buttons" drop-down list of the "Analyze" tab, which is available when a PivotChart is selected.
         */
        setShowAxisFieldButtons(showAxisFieldButtons: boolean): void;

        /**
         * Specifies whether to display the legend field buttons on a PivotChart.
         */
        getShowLegendFieldButtons(): boolean;

        /**
         * Specifies whether to display the legend field buttons on a PivotChart.
         */
        setShowLegendFieldButtons(showLegendFieldButtons: boolean): void;

        /**
         * Specifies whether to display the report filter field buttons on a PivotChart.
         */
        getShowReportFilterFieldButtons(): boolean;

        /**
         * Specifies whether to display the report filter field buttons on a PivotChart.
         */
        setShowReportFilterFieldButtons(
            showReportFilterFieldButtons: boolean
        ): void;

        /**
         * Specifies whether to display the show value field buttons on a PivotChart.
         */
        getShowValueFieldButtons(): boolean;

        /**
         * Specifies whether to display the show value field buttons on a PivotChart.
         */
        setShowValueFieldButtons(showValueFieldButtons: boolean): void;
    }

    /**
     * Encapsulates the format properties for the overall chart area.
     */
    interface ChartAreaFormat {
        /**
         * Represents the border format of chart area, which includes color, linestyle, and weight.
         */
        getBorder(): ChartBorder;

        /**
         * Specifies the color scheme of the chart.
         */
        getColorScheme(): ChartColorScheme;

        /**
         * Specifies the color scheme of the chart.
         */
        setColorScheme(colorScheme: ChartColorScheme): void;

        /**
         * Represents the fill format of an object, which includes background formatting information.
         */
        getFill(): ChartFill;

        /**
         * Represents the font attributes (font name, font size, color, etc.) for the current object.
         */
        getFont(): ChartFont;

        /**
         * Specifies if the chart area of the chart has rounded corners.
         */
        getRoundedCorners(): boolean;

        /**
         * Specifies if the chart area of the chart has rounded corners.
         */
        setRoundedCorners(roundedCorners: boolean): void;
    }

    /**
     * Represents a series in a chart.
     */
    interface ChartSeries {
        /**
         * Specifies the group for the specified series.
         */
        getAxisGroup(): ChartAxisGroup;

        /**
         * Specifies the group for the specified series.
         */
        setAxisGroup(axisGroup: ChartAxisGroup): void;

        /**
         * Encapsulates the bin options for histogram charts and pareto charts.
         */
        getBinOptions(): ChartBinOptions;

        /**
         * Encapsulates the options for the box and whisker charts.
         */
        getBoxwhiskerOptions(): ChartBoxwhiskerOptions;

        /**
         * This can be an integer value from 0 (zero) to 300, representing the percentage of the default size. This property only applies to bubble charts.
         */
        getBubbleScale(): number;

        /**
         * This can be an integer value from 0 (zero) to 300, representing the percentage of the default size. This property only applies to bubble charts.
         */
        setBubbleScale(bubbleScale: number): void;

        /**
         * Represents the chart type of a series. See `ExcelScript.ChartType` for details.
         */
        getChartType(): ChartType;

        /**
         * Represents the chart type of a series. See `ExcelScript.ChartType` for details.
         */
        setChartType(chartType: ChartType): void;

        /**
         * Represents a collection of all data labels in the series.
         */
        getDataLabels(): ChartDataLabels;

        /**
         * Represents the doughnut hole size of a chart series. Only valid on doughnut and doughnut exploded charts.
         * Throws an `InvalidArgument` error on invalid charts.
         */
        getDoughnutHoleSize(): number;

        /**
         * Represents the doughnut hole size of a chart series. Only valid on doughnut and doughnut exploded charts.
         * Throws an `InvalidArgument` error on invalid charts.
         */
        setDoughnutHoleSize(doughnutHoleSize: number): void;

        /**
         * Specifies the explosion value for a pie-chart or doughnut-chart slice. Returns 0 (zero) if there's no explosion (the tip of the slice is in the center of the pie).
         */
        getExplosion(): number;

        /**
         * Specifies the explosion value for a pie-chart or doughnut-chart slice. Returns 0 (zero) if there's no explosion (the tip of the slice is in the center of the pie).
         */
        setExplosion(explosion: number): void;

        /**
         * Specifies if the series is filtered. Not applicable for surface charts.
         */
        getFiltered(): boolean;

        /**
         * Specifies if the series is filtered. Not applicable for surface charts.
         */
        setFiltered(filtered: boolean): void;

        /**
         * Specifies the angle of the first pie-chart or doughnut-chart slice, in degrees (clockwise from vertical). Applies only to pie, 3-D pie, and doughnut charts. Can be a value from 0 through 360.
         */
        getFirstSliceAngle(): number;

        /**
         * Specifies the angle of the first pie-chart or doughnut-chart slice, in degrees (clockwise from vertical). Applies only to pie, 3-D pie, and doughnut charts. Can be a value from 0 through 360.
         */
        setFirstSliceAngle(firstSliceAngle: number): void;

        /**
         * Represents the formatting of a chart series, which includes fill and line formatting.
         */
        getFormat(): ChartSeriesFormat;

        /**
         * Represents the gap width of a chart series. Only valid on bar and column charts, as well as
         * specific classes of line and pie charts. Throws an invalid argument exception on invalid charts.
         */
        getGapWidth(): number;

        /**
         * Represents the gap width of a chart series. Only valid on bar and column charts, as well as
         * specific classes of line and pie charts. Throws an invalid argument exception on invalid charts.
         */
        setGapWidth(gapWidth: number): void;

        /**
         * Specifies the color for maximum value of a region map chart series.
         */
        getGradientMaximumColor(): string;

        /**
         * Specifies the color for maximum value of a region map chart series.
         */
        setGradientMaximumColor(gradientMaximumColor: string): void;

        /**
         * Specifies the type for maximum value of a region map chart series.
         */
        getGradientMaximumType(): ChartGradientStyleType;

        /**
         * Specifies the type for maximum value of a region map chart series.
         */
        setGradientMaximumType(
            gradientMaximumType: ChartGradientStyleType
        ): void;

        /**
         * Specifies the maximum value of a region map chart series.
         */
        getGradientMaximumValue(): number;

        /**
         * Specifies the maximum value of a region map chart series.
         */
        setGradientMaximumValue(gradientMaximumValue: number): void;

        /**
         * Specifies the color for the midpoint value of a region map chart series.
         */
        getGradientMidpointColor(): string;

        /**
         * Specifies the color for the midpoint value of a region map chart series.
         */
        setGradientMidpointColor(gradientMidpointColor: string): void;

        /**
         * Specifies the type for the midpoint value of a region map chart series.
         */
        getGradientMidpointType(): ChartGradientStyleType;

        /**
         * Specifies the type for the midpoint value of a region map chart series.
         */
        setGradientMidpointType(
            gradientMidpointType: ChartGradientStyleType
        ): void;

        /**
         * Specifies the midpoint value of a region map chart series.
         */
        getGradientMidpointValue(): number;

        /**
         * Specifies the midpoint value of a region map chart series.
         */
        setGradientMidpointValue(gradientMidpointValue: number): void;

        /**
         * Specifies the color for the minimum value of a region map chart series.
         */
        getGradientMinimumColor(): string;

        /**
         * Specifies the color for the minimum value of a region map chart series.
         */
        setGradientMinimumColor(gradientMinimumColor: string): void;

        /**
         * Specifies the type for the minimum value of a region map chart series.
         */
        getGradientMinimumType(): ChartGradientStyleType;

        /**
         * Specifies the type for the minimum value of a region map chart series.
         */
        setGradientMinimumType(
            gradientMinimumType: ChartGradientStyleType
        ): void;

        /**
         * Specifies the minimum value of a region map chart series.
         */
        getGradientMinimumValue(): number;

        /**
         * Specifies the minimum value of a region map chart series.
         */
        setGradientMinimumValue(gradientMinimumValue: number): void;

        /**
         * Specifies the series gradient style of a region map chart.
         */
        getGradientStyle(): ChartGradientStyle;

        /**
         * Specifies the series gradient style of a region map chart.
         */
        setGradientStyle(gradientStyle: ChartGradientStyle): void;

        /**
         * Specifies if the series has data labels.
         */
        getHasDataLabels(): boolean;

        /**
         * Specifies if the series has data labels.
         */
        setHasDataLabels(hasDataLabels: boolean): void;

        /**
         * Specifies the fill color for negative data points in a series.
         */
        getInvertColor(): string;

        /**
         * Specifies the fill color for negative data points in a series.
         */
        setInvertColor(invertColor: string): void;

        /**
         * True if Excel inverts the pattern in the item when it corresponds to a negative number.
         */
        getInvertIfNegative(): boolean;

        /**
         * True if Excel inverts the pattern in the item when it corresponds to a negative number.
         */
        setInvertIfNegative(invertIfNegative: boolean): void;

        /**
         * Encapsulates the options for a region map chart.
         */
        getMapOptions(): ChartMapOptions;

        /**
         * Specifies the marker background color of a chart series.
         */
        getMarkerBackgroundColor(): string;

        /**
         * Specifies the marker background color of a chart series.
         */
        setMarkerBackgroundColor(markerBackgroundColor: string): void;

        /**
         * Specifies the marker foreground color of a chart series.
         */
        getMarkerForegroundColor(): string;

        /**
         * Specifies the marker foreground color of a chart series.
         */
        setMarkerForegroundColor(markerForegroundColor: string): void;

        /**
         * Specifies the marker size of a chart series.
         */
        getMarkerSize(): number;

        /**
         * Specifies the marker size of a chart series.
         */
        setMarkerSize(markerSize: number): void;

        /**
         * Specifies the marker style of a chart series. See `ExcelScript.ChartMarkerStyle` for details.
         */
        getMarkerStyle(): ChartMarkerStyle;

        /**
         * Specifies the marker style of a chart series. See `ExcelScript.ChartMarkerStyle` for details.
         */
        setMarkerStyle(markerStyle: ChartMarkerStyle): void;

        /**
         * Specifies the name of a series in a chart. The name's length should not be greater than 255 characters.
         */
        getName(): string;

        /**
         * Specifies the name of a series in a chart. The name's length should not be greater than 255 characters.
         */
        setName(name: string): void;

        /**
         * Specifies how bars and columns are positioned. Can be a value between –100 and 100. Applies only to 2-D bar and 2-D column charts.
         */
        getOverlap(): number;

        /**
         * Specifies how bars and columns are positioned. Can be a value between –100 and 100. Applies only to 2-D bar and 2-D column charts.
         */
        setOverlap(overlap: number): void;

        /**
         * Specifies the series parent label strategy area for a treemap chart.
         */
        getParentLabelStrategy(): ChartParentLabelStrategy;

        /**
         * Specifies the series parent label strategy area for a treemap chart.
         */
        setParentLabelStrategy(
            parentLabelStrategy: ChartParentLabelStrategy
        ): void;

        /**
         * Specifies the plot order of a chart series within the chart group.
         */
        getPlotOrder(): number;

        /**
         * Specifies the plot order of a chart series within the chart group.
         */
        setPlotOrder(plotOrder: number): void;

        /**
         * Specifies the size of the secondary section of either a pie-of-pie chart or a bar-of-pie chart, as a percentage of the size of the primary pie. Can be a value from 5 to 200.
         */
        getSecondPlotSize(): number;

        /**
         * Specifies the size of the secondary section of either a pie-of-pie chart or a bar-of-pie chart, as a percentage of the size of the primary pie. Can be a value from 5 to 200.
         */
        setSecondPlotSize(secondPlotSize: number): void;

        /**
         * Specifies whether connector lines are shown in waterfall charts.
         */
        getShowConnectorLines(): boolean;

        /**
         * Specifies whether connector lines are shown in waterfall charts.
         */
        setShowConnectorLines(showConnectorLines: boolean): void;

        /**
         * Specifies whether leader lines are displayed for each data label in the series.
         */
        getShowLeaderLines(): boolean;

        /**
         * Specifies whether leader lines are displayed for each data label in the series.
         */
        setShowLeaderLines(showLeaderLines: boolean): void;

        /**
         * Specifies if the series has a shadow.
         */
        getShowShadow(): boolean;

        /**
         * Specifies if the series has a shadow.
         */
        setShowShadow(showShadow: boolean): void;

        /**
         * Specifies if the series is smooth. Only applicable to line and scatter charts.
         */
        getSmooth(): boolean;

        /**
         * Specifies if the series is smooth. Only applicable to line and scatter charts.
         */
        setSmooth(smooth: boolean): void;

        /**
         * Specifies the way the two sections of either a pie-of-pie chart or a bar-of-pie chart are split.
         */
        getSplitType(): ChartSplitType;

        /**
         * Specifies the way the two sections of either a pie-of-pie chart or a bar-of-pie chart are split.
         */
        setSplitType(splitType: ChartSplitType): void;

        /**
         * Specifies the threshold value that separates two sections of either a pie-of-pie chart or a bar-of-pie chart.
         */
        getSplitValue(): number;

        /**
         * Specifies the threshold value that separates two sections of either a pie-of-pie chart or a bar-of-pie chart.
         */
        setSplitValue(splitValue: number): void;

        /**
         * True if Excel assigns a different color or pattern to each data marker. The chart must contain only one series.
         */
        getVaryByCategories(): boolean;

        /**
         * True if Excel assigns a different color or pattern to each data marker. The chart must contain only one series.
         */
        setVaryByCategories(varyByCategories: boolean): void;

        /**
         * Represents the error bar object of a chart series.
         */
        getXErrorBars(): ChartErrorBars;

        /**
         * Represents the error bar object of a chart series.
         */
        getYErrorBars(): ChartErrorBars;

        /**
         * Deletes the chart series.
         */
        delete(): void;

        /**
         * Gets the values from a single dimension of the chart series. These could be either category values or data values, depending on the dimension specified and how the data is mapped for the chart series.
         * @param dimension The dimension of the axis where the data is from.
         */
        getDimensionValues(dimension: ChartSeriesDimension): string[];

        /**
         * Sets the bubble sizes for a chart series. Only works for bubble charts.
         * @param sourceData The `Range` object corresponding to the source data.
         */
        setBubbleSizes(sourceData: Range): void;

        /**
         * Sets the values for a chart series. For scatter charts, it refers to y-axis values.
         * @param sourceData The `Range` object corresponding to the source data.
         */
        setValues(sourceData: Range): void;

        /**
         * Sets the values of the x-axis for a chart series. Only works for scatter charts.
         * @param sourceData The `Range` object corresponding to the source data.
         */
        setXAxisValues(sourceData: Range): void;

        /**
         * Returns a collection of all points in the series.
         */
        getPoints(): ChartPoint[];

        /**
         * The collection of trendlines in the series.
         */
        getTrendlines(): ChartTrendline[];

        /**
         * Adds a new trendline to trendline collection.
         * @param type Specifies the trendline type. The default value is "Linear". See `ExcelScript.ChartTrendline` for details.
         */
        addChartTrendline(type?: ChartTrendlineType): ChartTrendline;

        /**
         * Gets a trendline object by index, which is the insertion order in the items array.
         * @param index Represents the insertion order in the items array.
         */
        getChartTrendline(index: number): ChartTrendline;
    }

    /**
     * Encapsulates the format properties for the chart series
     */
    interface ChartSeriesFormat {
        /**
         * Represents the fill format of a chart series, which includes background formatting information.
         */
        getFill(): ChartFill;

        /**
         * Represents line formatting.
         */
        getLine(): ChartLineFormat;
    }

    /**
     * Represents a point of a series in a chart.
     */
    interface ChartPoint {
        /**
         * Returns the data label of a chart point.
         */
        getDataLabel(): ChartDataLabel;

        /**
         * Encapsulates the format properties chart point.
         */
        getFormat(): ChartPointFormat;

        /**
         * Represents whether a data point has a data label. Not applicable for surface charts.
         */
        getHasDataLabel(): boolean;

        /**
         * Represents whether a data point has a data label. Not applicable for surface charts.
         */
        setHasDataLabel(hasDataLabel: boolean): void;

        /**
         * HTML color code representation of the marker background color of a data point (e.g., #FF0000 represents Red).
         */
        getMarkerBackgroundColor(): string;

        /**
         * HTML color code representation of the marker background color of a data point (e.g., #FF0000 represents Red).
         */
        setMarkerBackgroundColor(markerBackgroundColor: string): void;

        /**
         * HTML color code representation of the marker foreground color of a data point (e.g., #FF0000 represents Red).
         */
        getMarkerForegroundColor(): string;

        /**
         * HTML color code representation of the marker foreground color of a data point (e.g., #FF0000 represents Red).
         */
        setMarkerForegroundColor(markerForegroundColor: string): void;

        /**
         * Represents marker size of a data point.
         */
        getMarkerSize(): number;

        /**
         * Represents marker size of a data point.
         */
        setMarkerSize(markerSize: number): void;

        /**
         * Represents marker style of a chart data point. See `ExcelScript.ChartMarkerStyle` for details.
         */
        getMarkerStyle(): ChartMarkerStyle;

        /**
         * Represents marker style of a chart data point. See `ExcelScript.ChartMarkerStyle` for details.
         */
        setMarkerStyle(markerStyle: ChartMarkerStyle): void;

        /**
         * Returns the value of a chart point.
         */
        getValue(): number;
    }

    /**
     * Represents the formatting object for chart points.
     */
    interface ChartPointFormat {
        /**
         * Represents the border format of a chart data point, which includes color, style, and weight information.
         */
        getBorder(): ChartBorder;

        /**
         * Represents the fill format of a chart, which includes background formatting information.
         */
        getFill(): ChartFill;
    }

    /**
     * Represents the chart axes.
     */
    interface ChartAxes {
        /**
         * Represents the category axis in a chart.
         */
        getCategoryAxis(): ChartAxis;

        /**
         * Represents the series axis of a 3-D chart.
         */
        getSeriesAxis(): ChartAxis;

        /**
         * Represents the value axis in an axis.
         */
        getValueAxis(): ChartAxis;

        /**
         * Returns the specific axis identified by type and group.
         * @param type Specifies the axis type. See `ExcelScript.ChartAxisType` for details.
         * @param group Optional. Specifies the axis group. See `ExcelScript.ChartAxisGroup` for details.
         */
        getChartAxis(type: ChartAxisType, group?: ChartAxisGroup): ChartAxis;
    }

    /**
     * Represents a single axis in a chart.
     */
    interface ChartAxis {
        /**
         * Specifies the alignment for the specified axis tick label. See `ExcelScript.ChartTextHorizontalAlignment` for detail.
         */
        getAlignment(): ChartTickLabelAlignment;

        /**
         * Specifies the alignment for the specified axis tick label. See `ExcelScript.ChartTextHorizontalAlignment` for detail.
         */
        setAlignment(alignment: ChartTickLabelAlignment): void;

        /**
         * Specifies the group for the specified axis. See `ExcelScript.ChartAxisGroup` for details.
         */
        getAxisGroup(): ChartAxisGroup;

        /**
         * Specifies the base unit for the specified category axis.
         */
        getBaseTimeUnit(): ChartAxisTimeUnit;

        /**
         * Specifies the base unit for the specified category axis.
         */
        setBaseTimeUnit(baseTimeUnit: ChartAxisTimeUnit): void;

        /**
         * Specifies the category axis type.
         */
        getCategoryType(): ChartAxisCategoryType;

        /**
         * Specifies the category axis type.
         */
        setCategoryType(categoryType: ChartAxisCategoryType): void;

        /**
         * Specifies the custom axis display unit value. To set this property, please use the `SetCustomDisplayUnit(double)` method.
         */
        getCustomDisplayUnit(): number;

        /**
         * Represents the axis display unit. See `ExcelScript.ChartAxisDisplayUnit` for details.
         */
        getDisplayUnit(): ChartAxisDisplayUnit;

        /**
         * Represents the axis display unit. See `ExcelScript.ChartAxisDisplayUnit` for details.
         */
        setDisplayUnit(displayUnit: ChartAxisDisplayUnit): void;

        /**
         * Represents the formatting of a chart object, which includes line and font formatting.
         */
        getFormat(): ChartAxisFormat;

        /**
         * Specifies the height, in points, of the chart axis. Returns `null` if the axis is not visible.
         */
        getHeight(): number;

        /**
         * Specifies if the value axis crosses the category axis between categories.
         */
        getIsBetweenCategories(): boolean;

        /**
         * Specifies if the value axis crosses the category axis between categories.
         */
        setIsBetweenCategories(isBetweenCategories: boolean): void;

        /**
         * Specifies the distance, in points, from the left edge of the axis to the left of chart area. Returns `null` if the axis is not visible.
         */
        getLeft(): number;

        /**
         * Specifies if the number format is linked to the cells. If `true`, the number format will change in the labels when it changes in the cells.
         */
        getLinkNumberFormat(): boolean;

        /**
         * Specifies if the number format is linked to the cells. If `true`, the number format will change in the labels when it changes in the cells.
         */
        setLinkNumberFormat(linkNumberFormat: boolean): void;

        /**
         * Specifies the base of the logarithm when using logarithmic scales.
         */
        getLogBase(): number;

        /**
         * Specifies the base of the logarithm when using logarithmic scales.
         */
        setLogBase(logBase: number): void;

        /**
         * Returns an object that represents the major gridlines for the specified axis.
         */
        getMajorGridlines(): ChartGridlines;

        /**
         * Specifies the type of major tick mark for the specified axis. See `ExcelScript.ChartAxisTickMark` for details.
         */
        getMajorTickMark(): ChartAxisTickMark;

        /**
         * Specifies the type of major tick mark for the specified axis. See `ExcelScript.ChartAxisTickMark` for details.
         */
        setMajorTickMark(majorTickMark: ChartAxisTickMark): void;

        /**
         * Specifies the major unit scale value for the category axis when the `categoryType` property is set to `dateAxis`.
         */
        getMajorTimeUnitScale(): ChartAxisTimeUnit;

        /**
         * Specifies the major unit scale value for the category axis when the `categoryType` property is set to `dateAxis`.
         */
        setMajorTimeUnitScale(majorTimeUnitScale: ChartAxisTimeUnit): void;

        /**
         * Specifies the interval between two major tick marks.
         */
        getMajorUnit(): number;

        /**
         * Specifies the interval between two major tick marks.
         */
        setMajorUnit(majorUnit: number): void;

        /**
         * Specifies the maximum value on the value axis.
         */
        getMaximum(): number;

        /**
         * Specifies the maximum value on the value axis.
         */
        setMaximum(maximum: number): void;

        /**
         * Specifies the minimum value on the value axis.
         */
        getMinimum(): number;

        /**
         * Specifies the minimum value on the value axis.
         */
        setMinimum(minimum: number): void;

        /**
         * Returns an object that represents the minor gridlines for the specified axis.
         */
        getMinorGridlines(): ChartGridlines;

        /**
         * Specifies the type of minor tick mark for the specified axis. See `ExcelScript.ChartAxisTickMark` for details.
         */
        getMinorTickMark(): ChartAxisTickMark;

        /**
         * Specifies the type of minor tick mark for the specified axis. See `ExcelScript.ChartAxisTickMark` for details.
         */
        setMinorTickMark(minorTickMark: ChartAxisTickMark): void;

        /**
         * Specifies the minor unit scale value for the category axis when the `categoryType` property is set to `dateAxis`.
         */
        getMinorTimeUnitScale(): ChartAxisTimeUnit;

        /**
         * Specifies the minor unit scale value for the category axis when the `categoryType` property is set to `dateAxis`.
         */
        setMinorTimeUnitScale(minorTimeUnitScale: ChartAxisTimeUnit): void;

        /**
         * Specifies the interval between two minor tick marks.
         */
        getMinorUnit(): number;

        /**
         * Specifies the interval between two minor tick marks.
         */
        setMinorUnit(minorUnit: number): void;

        /**
         * Specifies if an axis is multilevel.
         */
        getMultiLevel(): boolean;

        /**
         * Specifies if an axis is multilevel.
         */
        setMultiLevel(multiLevel: boolean): void;

        /**
         * Specifies the format code for the axis tick label.
         */
        getNumberFormat(): string;

        /**
         * Specifies the format code for the axis tick label.
         */
        setNumberFormat(numberFormat: string): void;

        /**
         * Specifies the distance between the levels of labels, and the distance between the first level and the axis line. The value should be an integer from 0 to 1000.
         */
        getOffset(): number;

        /**
         * Specifies the distance between the levels of labels, and the distance between the first level and the axis line. The value should be an integer from 0 to 1000.
         */
        setOffset(offset: number): void;

        /**
         * Specifies the specified axis position where the other axis crosses. See `ExcelScript.ChartAxisPosition` for details.
         */
        getPosition(): ChartAxisPosition;

        /**
         * Specifies the specified axis position where the other axis crosses. See `ExcelScript.ChartAxisPosition` for details.
         */
        setPosition(position: ChartAxisPosition): void;

        /**
         * Specifies the axis position where the other axis crosses. You should use the `SetPositionAt(double)` method to set this property.
         */
        getPositionAt(): number;

        /**
         * Specifies if Excel plots data points from last to first.
         */
        getReversePlotOrder(): boolean;

        /**
         * Specifies if Excel plots data points from last to first.
         */
        setReversePlotOrder(reversePlotOrder: boolean): void;

        /**
         * Specifies the value axis scale type. See `ExcelScript.ChartAxisScaleType` for details.
         */
        getScaleType(): ChartAxisScaleType;

        /**
         * Specifies the value axis scale type. See `ExcelScript.ChartAxisScaleType` for details.
         */
        setScaleType(scaleType: ChartAxisScaleType): void;

        /**
         * Specifies if the axis display unit label is visible.
         */
        getShowDisplayUnitLabel(): boolean;

        /**
         * Specifies if the axis display unit label is visible.
         */
        setShowDisplayUnitLabel(showDisplayUnitLabel: boolean): void;

        /**
         * Specifies the angle to which the text is oriented for the chart axis tick label. The value should either be an integer from -90 to 90 or the integer 180 for vertically-oriented text.
         */
        getTextOrientation(): number;

        /**
         * Specifies the angle to which the text is oriented for the chart axis tick label. The value should either be an integer from -90 to 90 or the integer 180 for vertically-oriented text.
         */
        setTextOrientation(textOrientation: number): void;

        /**
         * Specifies the position of tick-mark labels on the specified axis. See `ExcelScript.ChartAxisTickLabelPosition` for details.
         */
        getTickLabelPosition(): ChartAxisTickLabelPosition;

        /**
         * Specifies the position of tick-mark labels on the specified axis. See `ExcelScript.ChartAxisTickLabelPosition` for details.
         */
        setTickLabelPosition(
            tickLabelPosition: ChartAxisTickLabelPosition
        ): void;

        /**
         * Specifies the number of categories or series between tick-mark labels. Can be a value from 1 through 31999.
         */
        getTickLabelSpacing(): number;

        /**
         * Specifies the number of categories or series between tick-mark labels. Can be a value from 1 through 31999.
         */
        setTickLabelSpacing(tickLabelSpacing: number): void;

        /**
         * Specifies the number of categories or series between tick marks.
         */
        getTickMarkSpacing(): number;

        /**
         * Specifies the number of categories or series between tick marks.
         */
        setTickMarkSpacing(tickMarkSpacing: number): void;

        /**
         * Represents the axis title.
         */
        getTitle(): ChartAxisTitle;

        /**
         * Specifies the distance, in points, from the top edge of the axis to the top of chart area. Returns `null` if the axis is not visible.
         */
        getTop(): number;

        /**
         * Specifies the axis type. See `ExcelScript.ChartAxisType` for details.
         */
        getType(): ChartAxisType;

        /**
         * Specifies if the axis is visible.
         */
        getVisible(): boolean;

        /**
         * Specifies if the axis is visible.
         */
        setVisible(visible: boolean): void;

        /**
         * Specifies the width, in points, of the chart axis. Returns `null` if the axis is not visible.
         */
        getWidth(): number;

        /**
         * Sets all the category names for the specified axis.
         * @param sourceData The `Range` object corresponding to the source data.
         */
        setCategoryNames(sourceData: Range): void;

        /**
         * Sets the axis display unit to a custom value.
         * @param value Custom value of the display unit.
         */
        setCustomDisplayUnit(value: number): void;

        /**
         * Sets the specified axis position where the other axis crosses.
         * @param value Custom value of the crossing point.
         */
        setPositionAt(value: number): void;
    }

    /**
     * Encapsulates the format properties for the chart axis.
     */
    interface ChartAxisFormat {
        /**
         * Specifies chart fill formatting.
         */
        getFill(): ChartFill;

        /**
         * Specifies the font attributes (font name, font size, color, etc.) for a chart axis element.
         */
        getFont(): ChartFont;

        /**
         * Specifies chart line formatting.
         */
        getLine(): ChartLineFormat;
    }

    /**
     * Represents the title of a chart axis.
     */
    interface ChartAxisTitle {
        /**
         * Specifies the formatting of the chart axis title.
         */
        getFormat(): ChartAxisTitleFormat;

        /**
         * Specifies the axis title.
         */
        getText(): string;

        /**
         * Specifies the axis title.
         */
        setText(text: string): void;

        /**
         * Specifies the angle to which the text is oriented for the chart axis title. The value should either be an integer from -90 to 90 or the integer 180 for vertically-oriented text.
         */
        getTextOrientation(): number;

        /**
         * Specifies the angle to which the text is oriented for the chart axis title. The value should either be an integer from -90 to 90 or the integer 180 for vertically-oriented text.
         */
        setTextOrientation(textOrientation: number): void;

        /**
         * Specifies if the axis title is visibile.
         */
        getVisible(): boolean;

        /**
         * Specifies if the axis title is visibile.
         */
        setVisible(visible: boolean): void;

        /**
         * A string value that represents the formula of chart axis title using A1-style notation.
         * @param formula A string that represents the formula to set.
         */
        setFormula(formula: string): void;
    }

    /**
     * Represents the chart axis title formatting.
     */
    interface ChartAxisTitleFormat {
        /**
         * Specifies the chart axis title's border format, which includes color, linestyle, and weight.
         */
        getBorder(): ChartBorder;

        /**
         * Specifies the chart axis title's fill formatting.
         */
        getFill(): ChartFill;

        /**
         * Specifies the chart axis title's font attributes, such as font name, font size, or color, of the chart axis title object.
         */
        getFont(): ChartFont;
    }

    /**
     * Represents a collection of all the data labels on a chart point.
     */
    interface ChartDataLabels {
        /**
         * Specifies if data labels automatically generate appropriate text based on context.
         */
        getAutoText(): boolean;

        /**
         * Specifies if data labels automatically generate appropriate text based on context.
         */
        setAutoText(autoText: boolean): void;

        /**
         * Specifies the format of chart data labels, which includes fill and font formatting.
         */
        getFormat(): ChartDataLabelFormat;

        /**
         * Specifies the horizontal alignment for chart data label. See `ExcelScript.ChartTextHorizontalAlignment` for details.
         * This property is valid only when the `TextOrientation` of data label is 0.
         */
        getHorizontalAlignment(): ChartTextHorizontalAlignment;

        /**
         * Specifies the horizontal alignment for chart data label. See `ExcelScript.ChartTextHorizontalAlignment` for details.
         * This property is valid only when the `TextOrientation` of data label is 0.
         */
        setHorizontalAlignment(
            horizontalAlignment: ChartTextHorizontalAlignment
        ): void;

        /**
         * Specifies if the number format is linked to the cells. If `true`, the number format will change in the labels when it changes in the cells.
         */
        getLinkNumberFormat(): boolean;

        /**
         * Specifies if the number format is linked to the cells. If `true`, the number format will change in the labels when it changes in the cells.
         */
        setLinkNumberFormat(linkNumberFormat: boolean): void;

        /**
         * Specifies the format code for data labels.
         */
        getNumberFormat(): string;

        /**
         * Specifies the format code for data labels.
         */
        setNumberFormat(numberFormat: string): void;

        /**
         * Value that represents the position of the data label. See `ExcelScript.ChartDataLabelPosition` for details.
         */
        getPosition(): ChartDataLabelPosition;

        /**
         * Value that represents the position of the data label. See `ExcelScript.ChartDataLabelPosition` for details.
         */
        setPosition(position: ChartDataLabelPosition): void;

        /**
         * String representing the separator used for the data labels on a chart.
         */
        getSeparator(): string;

        /**
         * String representing the separator used for the data labels on a chart.
         */
        setSeparator(separator: string): void;

        /**
         * Specifies if the data label bubble size is visible.
         */
        getShowBubbleSize(): boolean;

        /**
         * Specifies if the data label bubble size is visible.
         */
        setShowBubbleSize(showBubbleSize: boolean): void;

        /**
         * Specifies if the data label category name is visible.
         */
        getShowCategoryName(): boolean;

        /**
         * Specifies if the data label category name is visible.
         */
        setShowCategoryName(showCategoryName: boolean): void;

        /**
         * Specifies if the data label legend key is visible.
         */
        getShowLegendKey(): boolean;

        /**
         * Specifies if the data label legend key is visible.
         */
        setShowLegendKey(showLegendKey: boolean): void;

        /**
         * Specifies if the data label percentage is visible.
         */
        getShowPercentage(): boolean;

        /**
         * Specifies if the data label percentage is visible.
         */
        setShowPercentage(showPercentage: boolean): void;

        /**
         * Specifies if the data label series name is visible.
         */
        getShowSeriesName(): boolean;

        /**
         * Specifies if the data label series name is visible.
         */
        setShowSeriesName(showSeriesName: boolean): void;

        /**
         * Specifies if the data label value is visible.
         */
        getShowValue(): boolean;

        /**
         * Specifies if the data label value is visible.
         */
        setShowValue(showValue: boolean): void;

        /**
         * Represents the angle to which the text is oriented for data labels. The value should either be an integer from -90 to 90 or the integer 180 for vertically-oriented text.
         */
        getTextOrientation(): number;

        /**
         * Represents the angle to which the text is oriented for data labels. The value should either be an integer from -90 to 90 or the integer 180 for vertically-oriented text.
         */
        setTextOrientation(textOrientation: number): void;

        /**
         * Represents the vertical alignment of chart data label. See `ExcelScript.ChartTextVerticalAlignment` for details.
         * This property is valid only when `TextOrientation` of the data label is -90, 90, or 180.
         */
        getVerticalAlignment(): ChartTextVerticalAlignment;

        /**
         * Represents the vertical alignment of chart data label. See `ExcelScript.ChartTextVerticalAlignment` for details.
         * This property is valid only when `TextOrientation` of the data label is -90, 90, or 180.
         */
        setVerticalAlignment(
            verticalAlignment: ChartTextVerticalAlignment
        ): void;
    }

    /**
     * Represents the data label of a chart point.
     */
    interface ChartDataLabel {
        /**
         * Specifies if the data label automatically generates appropriate text based on context.
         */
        getAutoText(): boolean;

        /**
         * Specifies if the data label automatically generates appropriate text based on context.
         */
        setAutoText(autoText: boolean): void;

        /**
         * Represents the format of chart data label.
         */
        getFormat(): ChartDataLabelFormat;

        /**
         * String value that represents the formula of chart data label using A1-style notation.
         */
        getFormula(): string;

        /**
         * String value that represents the formula of chart data label using A1-style notation.
         */
        setFormula(formula: string): void;

        /**
         * Returns the height, in points, of the chart data label. Value is `null` if the chart data label is not visible.
         */
        getHeight(): number;

        /**
         * Represents the horizontal alignment for chart data label. See `ExcelScript.ChartTextHorizontalAlignment` for details.
         * This property is valid only when `TextOrientation` of data label is -90, 90, or 180.
         */
        getHorizontalAlignment(): ChartTextHorizontalAlignment;

        /**
         * Represents the horizontal alignment for chart data label. See `ExcelScript.ChartTextHorizontalAlignment` for details.
         * This property is valid only when `TextOrientation` of data label is -90, 90, or 180.
         */
        setHorizontalAlignment(
            horizontalAlignment: ChartTextHorizontalAlignment
        ): void;

        /**
         * Represents the distance, in points, from the left edge of chart data label to the left edge of chart area. Value is `null` if the chart data label is not visible.
         */
        getLeft(): number;

        /**
         * Represents the distance, in points, from the left edge of chart data label to the left edge of chart area. Value is `null` if the chart data label is not visible.
         */
        setLeft(left: number): void;

        /**
         * Specifies if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).
         */
        getLinkNumberFormat(): boolean;

        /**
         * Specifies if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).
         */
        setLinkNumberFormat(linkNumberFormat: boolean): void;

        /**
         * String value that represents the format code for data label.
         */
        getNumberFormat(): string;

        /**
         * String value that represents the format code for data label.
         */
        setNumberFormat(numberFormat: string): void;

        /**
         * Value that represents the position of the data label. See `ExcelScript.ChartDataLabelPosition` for details.
         */
        getPosition(): ChartDataLabelPosition;

        /**
         * Value that represents the position of the data label. See `ExcelScript.ChartDataLabelPosition` for details.
         */
        setPosition(position: ChartDataLabelPosition): void;

        /**
         * String representing the separator used for the data label on a chart.
         */
        getSeparator(): string;

        /**
         * String representing the separator used for the data label on a chart.
         */
        setSeparator(separator: string): void;

        /**
         * Specifies if the data label bubble size is visible.
         */
        getShowBubbleSize(): boolean;

        /**
         * Specifies if the data label bubble size is visible.
         */
        setShowBubbleSize(showBubbleSize: boolean): void;

        /**
         * Specifies if the data label category name is visible.
         */
        getShowCategoryName(): boolean;

        /**
         * Specifies if the data label category name is visible.
         */
        setShowCategoryName(showCategoryName: boolean): void;

        /**
         * Specifies if the data label legend key is visible.
         */
        getShowLegendKey(): boolean;

        /**
         * Specifies if the data label legend key is visible.
         */
        setShowLegendKey(showLegendKey: boolean): void;

        /**
         * Specifies if the data label percentage is visible.
         */
        getShowPercentage(): boolean;

        /**
         * Specifies if the data label percentage is visible.
         */
        setShowPercentage(showPercentage: boolean): void;

        /**
         * Specifies if the data label series name is visible.
         */
        getShowSeriesName(): boolean;

        /**
         * Specifies if the data label series name is visible.
         */
        setShowSeriesName(showSeriesName: boolean): void;

        /**
         * Specifies if the data label value is visible.
         */
        getShowValue(): boolean;

        /**
         * Specifies if the data label value is visible.
         */
        setShowValue(showValue: boolean): void;

        /**
         * String representing the text of the data label on a chart.
         */
        getText(): string;

        /**
         * String representing the text of the data label on a chart.
         */
        setText(text: string): void;

        /**
         * Represents the angle to which the text is oriented for the chart data label. The value should either be an integer from -90 to 90 or the integer 180 for vertically-oriented text.
         */
        getTextOrientation(): number;

        /**
         * Represents the angle to which the text is oriented for the chart data label. The value should either be an integer from -90 to 90 or the integer 180 for vertically-oriented text.
         */
        setTextOrientation(textOrientation: number): void;

        /**
         * Represents the distance, in points, from the top edge of chart data label to the top of chart area. Value is `null` if the chart data label is not visible.
         */
        getTop(): number;

        /**
         * Represents the distance, in points, from the top edge of chart data label to the top of chart area. Value is `null` if the chart data label is not visible.
         */
        setTop(top: number): void;

        /**
         * Represents the vertical alignment of chart data label. See `ExcelScript.ChartTextVerticalAlignment` for details.
         * This property is valid only when `TextOrientation` of data label is 0.
         */
        getVerticalAlignment(): ChartTextVerticalAlignment;

        /**
         * Represents the vertical alignment of chart data label. See `ExcelScript.ChartTextVerticalAlignment` for details.
         * This property is valid only when `TextOrientation` of data label is 0.
         */
        setVerticalAlignment(
            verticalAlignment: ChartTextVerticalAlignment
        ): void;

        /**
         * Returns the width, in points, of the chart data label. Value is `null` if the chart data label is not visible.
         */
        getWidth(): number;
    }

    /**
     * Encapsulates the format properties for the chart data labels.
     */
    interface ChartDataLabelFormat {
        /**
         * Represents the border format, which includes color, linestyle, and weight.
         */
        getBorder(): ChartBorder;

        /**
         * Represents the fill format of the current chart data label.
         */
        getFill(): ChartFill;

        /**
         * Represents the font attributes (such as font name, font size, and color) for a chart data label.
         */
        getFont(): ChartFont;
    }

    /**
     * This object represents the attributes for a chart's error bars.
     */
    interface ChartErrorBars {
        /**
         * Specifies if error bars have an end style cap.
         */
        getEndStyleCap(): boolean;

        /**
         * Specifies if error bars have an end style cap.
         */
        setEndStyleCap(endStyleCap: boolean): void;

        /**
         * Specifies the formatting type of the error bars.
         */
        getFormat(): ChartErrorBarsFormat;

        /**
         * Specifies which parts of the error bars to include.
         */
        getInclude(): ChartErrorBarsInclude;

        /**
         * Specifies which parts of the error bars to include.
         */
        setInclude(include: ChartErrorBarsInclude): void;

        /**
         * The type of range marked by the error bars.
         */
        getType(): ChartErrorBarsType;

        /**
         * The type of range marked by the error bars.
         */
        setType(type: ChartErrorBarsType): void;

        /**
         * Specifies whether the error bars are displayed.
         */
        getVisible(): boolean;

        /**
         * Specifies whether the error bars are displayed.
         */
        setVisible(visible: boolean): void;
    }

    /**
     * Encapsulates the format properties for chart error bars.
     */
    interface ChartErrorBarsFormat {
        /**
         * Represents the chart line formatting.
         */
        getLine(): ChartLineFormat;
    }

    /**
     * Represents major or minor gridlines on a chart axis.
     */
    interface ChartGridlines {
        /**
         * Represents the formatting of chart gridlines.
         */
        getFormat(): ChartGridlinesFormat;

        /**
         * Specifies if the axis gridlines are visible.
         */
        getVisible(): boolean;

        /**
         * Specifies if the axis gridlines are visible.
         */
        setVisible(visible: boolean): void;
    }

    /**
     * Encapsulates the format properties for chart gridlines.
     */
    interface ChartGridlinesFormat {
        /**
         * Represents chart line formatting.
         */
        getLine(): ChartLineFormat;
    }

    /**
     * Represents the legend in a chart.
     */
    interface ChartLegend {
        /**
         * Represents the formatting of a chart legend, which includes fill and font formatting.
         */
        getFormat(): ChartLegendFormat;

        /**
         * Specifies the height, in points, of the legend on the chart. Value is `null` if the legend is not visible.
         */
        getHeight(): number;

        /**
         * Specifies the height, in points, of the legend on the chart. Value is `null` if the legend is not visible.
         */
        setHeight(height: number): void;

        /**
         * Specifies the left value, in points, of the legend on the chart. Value is `null` if the legend is not visible.
         */
        getLeft(): number;

        /**
         * Specifies the left value, in points, of the legend on the chart. Value is `null` if the legend is not visible.
         */
        setLeft(left: number): void;

        /**
         * Specifies if the chart legend should overlap with the main body of the chart.
         */
        getOverlay(): boolean;

        /**
         * Specifies if the chart legend should overlap with the main body of the chart.
         */
        setOverlay(overlay: boolean): void;

        /**
         * Specifies the position of the legend on the chart. See `ExcelScript.ChartLegendPosition` for details.
         */
        getPosition(): ChartLegendPosition;

        /**
         * Specifies the position of the legend on the chart. See `ExcelScript.ChartLegendPosition` for details.
         */
        setPosition(position: ChartLegendPosition): void;

        /**
         * Specifies if the legend has a shadow on the chart.
         */
        getShowShadow(): boolean;

        /**
         * Specifies if the legend has a shadow on the chart.
         */
        setShowShadow(showShadow: boolean): void;

        /**
         * Specifies the top of a chart legend.
         */
        getTop(): number;

        /**
         * Specifies the top of a chart legend.
         */
        setTop(top: number): void;

        /**
         * Specifies if the chart legend is visible.
         */
        getVisible(): boolean;

        /**
         * Specifies if the chart legend is visible.
         */
        setVisible(visible: boolean): void;

        /**
         * Specifies the width, in points, of the legend on the chart. Value is `null` if the legend is not visible.
         */
        getWidth(): number;

        /**
         * Specifies the width, in points, of the legend on the chart. Value is `null` if the legend is not visible.
         */
        setWidth(width: number): void;

        /**
         * Represents a collection of legendEntries in the legend.
         */
        getLegendEntries(): ChartLegendEntry[];
    }

    /**
     * Represents the legend entry in `legendEntryCollection`.
     */
    interface ChartLegendEntry {
        /**
         * Specifies the height of the legend entry on the chart legend.
         */
        getHeight(): number;

        /**
         * Specifies the index of the legend entry in the chart legend.
         */
        getIndex(): number;

        /**
         * Specifies the left value of a chart legend entry.
         */
        getLeft(): number;

        /**
         * Specifies the top of a chart legend entry.
         */
        getTop(): number;

        /**
         * Represents the visibility of a chart legend entry.
         */
        getVisible(): boolean;

        /**
         * Represents the visibility of a chart legend entry.
         */
        setVisible(visible: boolean): void;

        /**
         * Represents the width of the legend entry on the chart Legend.
         */
        getWidth(): number;
    }

    /**
     * Encapsulates the format properties of a chart legend.
     */
    interface ChartLegendFormat {
        /**
         * Represents the border format, which includes color, linestyle, and weight.
         */
        getBorder(): ChartBorder;

        /**
         * Represents the fill format of an object, which includes background formatting information.
         */
        getFill(): ChartFill;

        /**
         * Represents the font attributes such as font name, font size, and color of a chart legend.
         */
        getFont(): ChartFont;
    }

    /**
     * Encapsulates the properties for a region map chart.
     */
    interface ChartMapOptions {
        /**
         * Specifies the series map labels strategy of a region map chart.
         */
        getLabelStrategy(): ChartMapLabelStrategy;

        /**
         * Specifies the series map labels strategy of a region map chart.
         */
        setLabelStrategy(labelStrategy: ChartMapLabelStrategy): void;

        /**
         * Specifies the series mapping level of a region map chart.
         */
        getLevel(): ChartMapAreaLevel;

        /**
         * Specifies the series mapping level of a region map chart.
         */
        setLevel(level: ChartMapAreaLevel): void;

        /**
         * Specifies the series projection type of a region map chart.
         */
        getProjectionType(): ChartMapProjectionType;

        /**
         * Specifies the series projection type of a region map chart.
         */
        setProjectionType(projectionType: ChartMapProjectionType): void;
    }

    /**
     * Represents a chart title object of a chart.
     */
    interface ChartTitle {
        /**
         * Represents the formatting of a chart title, which includes fill and font formatting.
         */
        getFormat(): ChartTitleFormat;

        /**
         * Returns the height, in points, of the chart title. Value is `null` if the chart title is not visible.
         */
        getHeight(): number;

        /**
         * Specifies the horizontal alignment for chart title.
         */
        getHorizontalAlignment(): ChartTextHorizontalAlignment;

        /**
         * Specifies the horizontal alignment for chart title.
         */
        setHorizontalAlignment(
            horizontalAlignment: ChartTextHorizontalAlignment
        ): void;

        /**
         * Specifies the distance, in points, from the left edge of chart title to the left edge of chart area. Value is `null` if the chart title is not visible.
         */
        getLeft(): number;

        /**
         * Specifies the distance, in points, from the left edge of chart title to the left edge of chart area. Value is `null` if the chart title is not visible.
         */
        setLeft(left: number): void;

        /**
         * Specifies if the chart title will overlay the chart.
         */
        getOverlay(): boolean;

        /**
         * Specifies if the chart title will overlay the chart.
         */
        setOverlay(overlay: boolean): void;

        /**
         * Represents the position of chart title. See `ExcelScript.ChartTitlePosition` for details.
         */
        getPosition(): ChartTitlePosition;

        /**
         * Represents the position of chart title. See `ExcelScript.ChartTitlePosition` for details.
         */
        setPosition(position: ChartTitlePosition): void;

        /**
         * Represents a boolean value that determines if the chart title has a shadow.
         */
        getShowShadow(): boolean;

        /**
         * Represents a boolean value that determines if the chart title has a shadow.
         */
        setShowShadow(showShadow: boolean): void;

        /**
         * Specifies the chart's title text.
         */
        getText(): string;

        /**
         * Specifies the chart's title text.
         */
        setText(text: string): void;

        /**
         * Specifies the angle to which the text is oriented for the chart title. The value should either be an integer from -90 to 90 or the integer 180 for vertically-oriented text.
         */
        getTextOrientation(): number;

        /**
         * Specifies the angle to which the text is oriented for the chart title. The value should either be an integer from -90 to 90 or the integer 180 for vertically-oriented text.
         */
        setTextOrientation(textOrientation: number): void;

        /**
         * Specifies the distance, in points, from the top edge of chart title to the top of chart area. Value is `null` if the chart title is not visible.
         */
        getTop(): number;

        /**
         * Specifies the distance, in points, from the top edge of chart title to the top of chart area. Value is `null` if the chart title is not visible.
         */
        setTop(top: number): void;

        /**
         * Specifies the vertical alignment of chart title. See `ExcelScript.ChartTextVerticalAlignment` for details.
         */
        getVerticalAlignment(): ChartTextVerticalAlignment;

        /**
         * Specifies the vertical alignment of chart title. See `ExcelScript.ChartTextVerticalAlignment` for details.
         */
        setVerticalAlignment(
            verticalAlignment: ChartTextVerticalAlignment
        ): void;

        /**
         * Specifies if the chart title is visibile.
         */
        getVisible(): boolean;

        /**
         * Specifies if the chart title is visibile.
         */
        setVisible(visible: boolean): void;

        /**
         * Specifies the width, in points, of the chart title. Value is `null` if the chart title is not visible.
         */
        getWidth(): number;

        /**
         * Get the substring of a chart title. Line break '\n' counts one character.
         * @param start Start position of substring to be retrieved. Zero-indexed.
         * @param length Length of the substring to be retrieved.
         */
        getSubstring(start: number, length: number): ChartFormatString;

        /**
         * Sets a string value that represents the formula of chart title using A1-style notation.
         * @param formula A string that represents the formula to set.
         */
        setFormula(formula: string): void;
    }

    /**
     * Represents the substring in chart related objects that contain text, like a `ChartTitle` object or `ChartAxisTitle` object.
     */
    interface ChartFormatString {
        /**
         * Represents the font attributes, such as font name, font size, and color of a chart characters object.
         */
        getFont(): ChartFont;
    }

    /**
     * Provides access to the formatting options for a chart title.
     */
    interface ChartTitleFormat {
        /**
         * Represents the border format of chart title, which includes color, linestyle, and weight.
         */
        getBorder(): ChartBorder;

        /**
         * Represents the fill format of an object, which includes background formatting information.
         */
        getFill(): ChartFill;

        /**
         * Represents the font attributes (such as font name, font size, and color) for an object.
         */
        getFont(): ChartFont;
    }

    /**
     * Represents the fill formatting for a chart element.
     */
    interface ChartFill {
        /**
         * Clears the fill color of a chart element.
         */
        clear(): void;

        /**
         * Sets the fill formatting of a chart element to a uniform color.
         * @param color HTML color code representing the color of the background, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
         */
        setSolidColor(color: string): void;
    }

    /**
     * Represents the border formatting of a chart element.
     */
    interface ChartBorder {
        /**
         * HTML color code representing the color of borders in the chart.
         */
        getColor(): string;

        /**
         * HTML color code representing the color of borders in the chart.
         */
        setColor(color: string): void;

        /**
         * Represents the line style of the border. See `ExcelScript.ChartLineStyle` for details.
         */
        getLineStyle(): ChartLineStyle;

        /**
         * Represents the line style of the border. See `ExcelScript.ChartLineStyle` for details.
         */
        setLineStyle(lineStyle: ChartLineStyle): void;

        /**
         * Represents weight of the border, in points.
         */
        getWeight(): number;

        /**
         * Represents weight of the border, in points.
         */
        setWeight(weight: number): void;

        /**
         * Clear the border format of a chart element.
         */
        clear(): void;
    }

    /**
     * Encapsulates the bin options for histogram charts and pareto charts.
     */
    interface ChartBinOptions {
        /**
         * Specifies if bin overflow is enabled in a histogram chart or pareto chart.
         */
        getAllowOverflow(): boolean;

        /**
         * Specifies if bin overflow is enabled in a histogram chart or pareto chart.
         */
        setAllowOverflow(allowOverflow: boolean): void;

        /**
         * Specifies if bin underflow is enabled in a histogram chart or pareto chart.
         */
        getAllowUnderflow(): boolean;

        /**
         * Specifies if bin underflow is enabled in a histogram chart or pareto chart.
         */
        setAllowUnderflow(allowUnderflow: boolean): void;

        /**
         * Specifies the bin count of a histogram chart or pareto chart.
         */
        getCount(): number;

        /**
         * Specifies the bin count of a histogram chart or pareto chart.
         */
        setCount(count: number): void;

        /**
         * Specifies the bin overflow value of a histogram chart or pareto chart.
         */
        getOverflowValue(): number;

        /**
         * Specifies the bin overflow value of a histogram chart or pareto chart.
         */
        setOverflowValue(overflowValue: number): void;

        /**
         * Specifies the bin's type for a histogram chart or pareto chart.
         */
        getType(): ChartBinType;

        /**
         * Specifies the bin's type for a histogram chart or pareto chart.
         */
        setType(type: ChartBinType): void;

        /**
         * Specifies the bin underflow value of a histogram chart or pareto chart.
         */
        getUnderflowValue(): number;

        /**
         * Specifies the bin underflow value of a histogram chart or pareto chart.
         */
        setUnderflowValue(underflowValue: number): void;

        /**
         * Specifies the bin width value of a histogram chart or pareto chart.
         */
        getWidth(): number;

        /**
         * Specifies the bin width value of a histogram chart or pareto chart.
         */
        setWidth(width: number): void;
    }

    /**
     * Represents the properties of a box and whisker chart.
     */
    interface ChartBoxwhiskerOptions {
        /**
         * Specifies if the quartile calculation type of a box and whisker chart.
         */
        getQuartileCalculation(): ChartBoxQuartileCalculation;

        /**
         * Specifies if the quartile calculation type of a box and whisker chart.
         */
        setQuartileCalculation(
            quartileCalculation: ChartBoxQuartileCalculation
        ): void;

        /**
         * Specifies if inner points are shown in a box and whisker chart.
         */
        getShowInnerPoints(): boolean;

        /**
         * Specifies if inner points are shown in a box and whisker chart.
         */
        setShowInnerPoints(showInnerPoints: boolean): void;

        /**
         * Specifies if the mean line is shown in a box and whisker chart.
         */
        getShowMeanLine(): boolean;

        /**
         * Specifies if the mean line is shown in a box and whisker chart.
         */
        setShowMeanLine(showMeanLine: boolean): void;

        /**
         * Specifies if the mean marker is shown in a box and whisker chart.
         */
        getShowMeanMarker(): boolean;

        /**
         * Specifies if the mean marker is shown in a box and whisker chart.
         */
        setShowMeanMarker(showMeanMarker: boolean): void;

        /**
         * Specifies if outlier points are shown in a box and whisker chart.
         */
        getShowOutlierPoints(): boolean;

        /**
         * Specifies if outlier points are shown in a box and whisker chart.
         */
        setShowOutlierPoints(showOutlierPoints: boolean): void;
    }

    /**
     * Encapsulates the formatting options for line elements.
     */
    interface ChartLineFormat {
        /**
         * HTML color code representing the color of lines in the chart.
         */
        getColor(): string;

        /**
         * HTML color code representing the color of lines in the chart.
         */
        setColor(color: string): void;

        /**
         * Represents the line style. See `ExcelScript.ChartLineStyle` for details.
         */
        getLineStyle(): ChartLineStyle;

        /**
         * Represents the line style. See `ExcelScript.ChartLineStyle` for details.
         */
        setLineStyle(lineStyle: ChartLineStyle): void;

        /**
         * Represents weight of the line, in points.
         */
        getWeight(): number;

        /**
         * Represents weight of the line, in points.
         */
        setWeight(weight: number): void;

        /**
         * Clears the line format of a chart element.
         */
        clear(): void;
    }

    /**
     * This object represents the font attributes (such as font name, font size, and color) for a chart object.
     */
    interface ChartFont {
        /**
         * Represents the bold status of font.
         */
        getBold(): boolean;

        /**
         * Represents the bold status of font.
         */
        setBold(bold: boolean): void;

        /**
         * HTML color code representation of the text color (e.g., #FF0000 represents Red).
         */
        getColor(): string;

        /**
         * HTML color code representation of the text color (e.g., #FF0000 represents Red).
         */
        setColor(color: string): void;

        /**
         * Represents the italic status of the font.
         */
        getItalic(): boolean;

        /**
         * Represents the italic status of the font.
         */
        setItalic(italic: boolean): void;

        /**
         * Font name (e.g., "Calibri")
         */
        getName(): string;

        /**
         * Font name (e.g., "Calibri")
         */
        setName(name: string): void;

        /**
         * Size of the font (e.g., 11)
         */
        getSize(): number;

        /**
         * Size of the font (e.g., 11)
         */
        setSize(size: number): void;

        /**
         * Type of underline applied to the font. See `ExcelScript.ChartUnderlineStyle` for details.
         */
        getUnderline(): ChartUnderlineStyle;

        /**
         * Type of underline applied to the font. See `ExcelScript.ChartUnderlineStyle` for details.
         */
        setUnderline(underline: ChartUnderlineStyle): void;
    }

    /**
     * This object represents the attributes for a chart trendline object.
     */
    interface ChartTrendline {
        /**
         * Represents the number of periods that the trendline extends backward.
         */
        getBackwardPeriod(): number;

        /**
         * Represents the number of periods that the trendline extends backward.
         */
        setBackwardPeriod(backwardPeriod: number): void;

        /**
         * Represents the formatting of a chart trendline.
         */
        getFormat(): ChartTrendlineFormat;

        /**
         * Represents the number of periods that the trendline extends forward.
         */
        getForwardPeriod(): number;

        /**
         * Represents the number of periods that the trendline extends forward.
         */
        setForwardPeriod(forwardPeriod: number): void;

        /**
         * Specifies the intercept value of the trendline.
         */
        getIntercept(): number;

        /**
         * Specifies the intercept value of the trendline.
         */
        setIntercept(intercept: number): void;

        /**
         * Represents the label of a chart trendline.
         */
        getLabel(): ChartTrendlineLabel;

        /**
         * Represents the period of a chart trendline. Only applicable to trendlines with the type `MovingAverage`.
         */
        getMovingAveragePeriod(): number;

        /**
         * Represents the period of a chart trendline. Only applicable to trendlines with the type `MovingAverage`.
         */
        setMovingAveragePeriod(movingAveragePeriod: number): void;

        /**
         * Represents the name of the trendline. Can be set to a string value, a `null` value represents automatic values. The returned value is always a string
         */
        getName(): string;

        /**
         * Represents the name of the trendline. Can be set to a string value, a `null` value represents automatic values. The returned value is always a string
         */
        setName(name: string): void;

        /**
         * Represents the order of a chart trendline. Only applicable to trendlines with the type `Polynomial`.
         */
        getPolynomialOrder(): number;

        /**
         * Represents the order of a chart trendline. Only applicable to trendlines with the type `Polynomial`.
         */
        setPolynomialOrder(polynomialOrder: number): void;

        /**
         * True if the equation for the trendline is displayed on the chart.
         */
        getShowEquation(): boolean;

        /**
         * True if the equation for the trendline is displayed on the chart.
         */
        setShowEquation(showEquation: boolean): void;

        /**
         * True if the r-squared value for the trendline is displayed on the chart.
         */
        getShowRSquared(): boolean;

        /**
         * True if the r-squared value for the trendline is displayed on the chart.
         */
        setShowRSquared(showRSquared: boolean): void;

        /**
         * Represents the type of a chart trendline.
         */
        getType(): ChartTrendlineType;

        /**
         * Represents the type of a chart trendline.
         */
        setType(type: ChartTrendlineType): void;

        /**
         * Delete the trendline object.
         */
        delete(): void;
    }

    /**
     * Represents the format properties for the chart trendline.
     */
    interface ChartTrendlineFormat {
        /**
         * Represents chart line formatting.
         */
        getLine(): ChartLineFormat;
    }

    /**
     * This object represents the attributes for a chart trendline label object.
     */
    interface ChartTrendlineLabel {
        /**
         * Specifies if the trendline label automatically generates appropriate text based on context.
         */
        getAutoText(): boolean;

        /**
         * Specifies if the trendline label automatically generates appropriate text based on context.
         */
        setAutoText(autoText: boolean): void;

        /**
         * The format of the chart trendline label.
         */
        getFormat(): ChartTrendlineLabelFormat;

        /**
         * String value that represents the formula of the chart trendline label using A1-style notation.
         */
        getFormula(): string;

        /**
         * String value that represents the formula of the chart trendline label using A1-style notation.
         */
        setFormula(formula: string): void;

        /**
         * Returns the height, in points, of the chart trendline label. Value is `null` if the chart trendline label is not visible.
         */
        getHeight(): number;

        /**
         * Represents the horizontal alignment of the chart trendline label. See `ExcelScript.ChartTextHorizontalAlignment` for details.
         * This property is valid only when `TextOrientation` of a trendline label is -90, 90, or 180.
         */
        getHorizontalAlignment(): ChartTextHorizontalAlignment;

        /**
         * Represents the horizontal alignment of the chart trendline label. See `ExcelScript.ChartTextHorizontalAlignment` for details.
         * This property is valid only when `TextOrientation` of a trendline label is -90, 90, or 180.
         */
        setHorizontalAlignment(
            horizontalAlignment: ChartTextHorizontalAlignment
        ): void;

        /**
         * Represents the distance, in points, from the left edge of the chart trendline label to the left edge of the chart area. Value is `null` if the chart trendline label is not visible.
         */
        getLeft(): number;

        /**
         * Represents the distance, in points, from the left edge of the chart trendline label to the left edge of the chart area. Value is `null` if the chart trendline label is not visible.
         */
        setLeft(left: number): void;

        /**
         * Specifies if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).
         */
        getLinkNumberFormat(): boolean;

        /**
         * Specifies if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).
         */
        setLinkNumberFormat(linkNumberFormat: boolean): void;

        /**
         * String value that represents the format code for the trendline label.
         */
        getNumberFormat(): string;

        /**
         * String value that represents the format code for the trendline label.
         */
        setNumberFormat(numberFormat: string): void;

        /**
         * String representing the text of the trendline label on a chart.
         */
        getText(): string;

        /**
         * String representing the text of the trendline label on a chart.
         */
        setText(text: string): void;

        /**
         * Represents the angle to which the text is oriented for the chart trendline label. The value should either be an integer from -90 to 90 or the integer 180 for vertically-oriented text.
         */
        getTextOrientation(): number;

        /**
         * Represents the angle to which the text is oriented for the chart trendline label. The value should either be an integer from -90 to 90 or the integer 180 for vertically-oriented text.
         */
        setTextOrientation(textOrientation: number): void;

        /**
         * Represents the distance, in points, from the top edge of the chart trendline label to the top of the chart area. Value is `null` if the chart trendline label is not visible.
         */
        getTop(): number;

        /**
         * Represents the distance, in points, from the top edge of the chart trendline label to the top of the chart area. Value is `null` if the chart trendline label is not visible.
         */
        setTop(top: number): void;

        /**
         * Represents the vertical alignment of the chart trendline label. See `ExcelScript.ChartTextVerticalAlignment` for details.
         * This property is valid only when `TextOrientation` of a trendline label is 0.
         */
        getVerticalAlignment(): ChartTextVerticalAlignment;

        /**
         * Represents the vertical alignment of the chart trendline label. See `ExcelScript.ChartTextVerticalAlignment` for details.
         * This property is valid only when `TextOrientation` of a trendline label is 0.
         */
        setVerticalAlignment(
            verticalAlignment: ChartTextVerticalAlignment
        ): void;

        /**
         * Returns the width, in points, of the chart trendline label. Value is `null` if the chart trendline label is not visible.
         */
        getWidth(): number;
    }

    /**
     * Encapsulates the format properties for the chart trendline label.
     */
    interface ChartTrendlineLabelFormat {
        /**
         * Specifies the border format, which includes color, linestyle, and weight.
         */
        getBorder(): ChartBorder;

        /**
         * Specifies the fill format of the current chart trendline label.
         */
        getFill(): ChartFill;

        /**
         * Specifies the font attributes (such as font name, font size, and color) for a chart trendline label.
         */
        getFont(): ChartFont;
    }

    /**
     * This object represents the attributes for a chart plot area.
     */
    interface ChartPlotArea {
        /**
         * Specifies the formatting of a chart plot area.
         */
        getFormat(): ChartPlotAreaFormat;

        /**
         * Specifies the height value of a plot area.
         */
        getHeight(): number;

        /**
         * Specifies the height value of a plot area.
         */
        setHeight(height: number): void;

        /**
         * Specifies the inside height value of a plot area.
         */
        getInsideHeight(): number;

        /**
         * Specifies the inside height value of a plot area.
         */
        setInsideHeight(insideHeight: number): void;

        /**
         * Specifies the inside left value of a plot area.
         */
        getInsideLeft(): number;

        /**
         * Specifies the inside left value of a plot area.
         */
        setInsideLeft(insideLeft: number): void;

        /**
         * Specifies the inside top value of a plot area.
         */
        getInsideTop(): number;

        /**
         * Specifies the inside top value of a plot area.
         */
        setInsideTop(insideTop: number): void;

        /**
         * Specifies the inside width value of a plot area.
         */
        getInsideWidth(): number;

        /**
         * Specifies the inside width value of a plot area.
         */
        setInsideWidth(insideWidth: number): void;

        /**
         * Specifies the left value of a plot area.
         */
        getLeft(): number;

        /**
         * Specifies the left value of a plot area.
         */
        setLeft(left: number): void;

        /**
         * Specifies the position of a plot area.
         */
        getPosition(): ChartPlotAreaPosition;

        /**
         * Specifies the position of a plot area.
         */
        setPosition(position: ChartPlotAreaPosition): void;

        /**
         * Specifies the top value of a plot area.
         */
        getTop(): number;

        /**
         * Specifies the top value of a plot area.
         */
        setTop(top: number): void;

        /**
         * Specifies the width value of a plot area.
         */
        getWidth(): number;

        /**
         * Specifies the width value of a plot area.
         */
        setWidth(width: number): void;
    }

    /**
     * Represents the format properties for a chart plot area.
     */
    interface ChartPlotAreaFormat {
        /**
         * Specifies the border attributes of a chart plot area.
         */
        getBorder(): ChartBorder;

        /**
         * Specifies the fill format of an object, which includes background formatting information.
         */
        getFill(): ChartFill;
    }

    /**
     * Manages sorting operations on `Range` objects.
     */
    interface RangeSort {
        /**
         * Perform a sort operation.
         * @param fields The list of conditions to sort on.
         * @param matchCase Optional. Whether to have the casing impact string ordering.
         * @param hasHeaders Optional. Whether the range has a header.
         * @param orientation Optional. Whether the operation is sorting rows or columns.
         * @param method Optional. The ordering method used for Chinese characters.
         */
        apply(
            fields: SortField[],
            matchCase?: boolean,
            hasHeaders?: boolean,
            orientation?: SortOrientation,
            method?: SortMethod
        ): void;
    }

    /**
     * Manages sorting operations on `Table` objects.
     */
    interface TableSort {
        /**
         * Specifies the current conditions used to last sort the table.
         */
        getFields(): SortField[];

        /**
         * Specifies if the casing impacts the last sort of the table.
         */
        getMatchCase(): boolean;

        /**
         * Represents the Chinese character ordering method last used to sort the table.
         */
        getMethod(): SortMethod;

        /**
         * Perform a sort operation.
         * @param fields The list of conditions to sort on.
         * @param matchCase Optional. Whether to have the casing impact string ordering.
         * @param method Optional. The ordering method used for Chinese characters.
         */
        apply(
            fields: SortField[],
            matchCase?: boolean,
            method?: SortMethod
        ): void;

        /**
         * Clears the sorting that is currently on the table. While this doesn't modify the table's ordering, it clears the state of the header buttons.
         */
        clear(): void;

        /**
         * Reapplies the current sorting parameters to the table.
         */
        reapply(): void;
    }

    /**
     * Manages the filtering of a table's column.
     */
    interface Filter {
        /**
         * The currently applied filter on the given column.
         */
        getCriteria(): FilterCriteria;

        /**
         * Apply the given filter criteria on the given column.
         * @param criteria The criteria to apply.
         */
        apply(criteria: FilterCriteria): void;

        /**
         * Apply a "Bottom Item" filter to the column for the given number of elements.
         * @param count The number of elements from the bottom to show.
         */
        applyBottomItemsFilter(count: number): void;

        /**
         * Apply a "Bottom Percent" filter to the column for the given percentage of elements.
         * @param percent The percentage of elements from the bottom to show.
         */
        applyBottomPercentFilter(percent: number): void;

        /**
         * Apply a "Cell Color" filter to the column for the given color.
         * @param color The background color of the cells to show.
         */
        applyCellColorFilter(color: string): void;

        /**
         * Apply an "Icon" filter to the column for the given criteria strings.
         * @param criteria1 The first criteria string.
         * @param criteria2 Optional. The second criteria string.
         * @param oper Optional. The operator that describes how the two criteria are joined.
         */
        applyCustomFilter(
            criteria1: string,
            criteria2?: string,
            oper?: FilterOperator
        ): void;

        /**
         * Apply a "Dynamic" filter to the column.
         * @param criteria The dynamic criteria to apply.
         */
        applyDynamicFilter(criteria: DynamicFilterCriteria): void;

        /**
         * Apply a "Font Color" filter to the column for the given color.
         * @param color The font color of the cells to show.
         */
        applyFontColorFilter(color: string): void;

        /**
         * Apply an "Icon" filter to the column for the given icon.
         * @param icon The icons of the cells to show.
         */
        applyIconFilter(icon: Icon): void;

        /**
         * Apply a "Top Item" filter to the column for the given number of elements.
         * @param count The number of elements from the top to show.
         */
        applyTopItemsFilter(count: number): void;

        /**
         * Apply a "Top Percent" filter to the column for the given percentage of elements.
         * @param percent The percentage of elements from the top to show.
         */
        applyTopPercentFilter(percent: number): void;

        /**
         * Apply a "Values" filter to the column for the given values.
         * @param values The list of values to show. This must be an array of strings or an array of `ExcelScript.FilterDateTime` objects.
         */
        applyValuesFilter(values: Array<string | FilterDatetime>): void;

        /**
         * Clear the filter on the given column.
         */
        clear(): void;
    }

    /**
     * Represents the `AutoFilter` object.
     * AutoFilter turns the values in Excel column into specific filters based on the cell contents.
     */
    interface AutoFilter {
        /**
         * An array that holds all the filter criteria in the autofiltered range.
         */
        getCriteria(): FilterCriteria[];

        /**
         * Specifies if the AutoFilter is enabled.
         */
        getEnabled(): boolean;

        /**
         * Specifies if the AutoFilter has filter criteria.
         */
        getIsDataFiltered(): boolean;

        /**
         * Applies the AutoFilter to a range. This filters the column if column index and filter criteria are specified.
         * @param range The range on which the AutoFilter will apply.
         * @param columnIndex The zero-based column index to which the AutoFilter is applied.
         * @param criteria The filter criteria.
         */
        apply(
            range: Range | string,
            columnIndex?: number,
            criteria?: FilterCriteria
        ): void;

        /**
         * Clears the filter criteria and sort state of the AutoFilter.
         */
        clearCriteria(): void;

        /**
         * Returns the `Range` object that represents the range to which the AutoFilter applies.
         * If there is no `Range` object associated with the AutoFilter, then this method returns `undefined`.
         */
        getRange(): Range;

        /**
         * Applies the specified AutoFilter object currently on the range.
         */
        reapply(): void;

        /**
         * Removes the AutoFilter for the range.
         */
        remove(): void;
    }

    /**
     * Provides information based on current system culture settings. This includes the culture names, number formatting, and other culturally dependent settings.
     */
    interface CultureInfo {
        /**
         * Defines the culturally appropriate format of displaying date and time. This is based on current system culture settings.
         */
        getDatetimeFormat(): DatetimeFormatInfo;

        /**
         * Gets the culture name in the format languagecode2-country/regioncode2 (e.g., "zh-cn" or "en-us"). This is based on current system settings.
         */
        getName(): string;

        /**
         * Defines the culturally appropriate format of displaying numbers. This is based on current system culture settings.
         */
        getNumberFormat(): NumberFormatInfo;
    }

    /**
     * Defines the culturally appropriate format of displaying numbers. This is based on current system culture settings.
     */
    interface NumberFormatInfo {
        /**
         * Gets the string used as the decimal separator for numeric values. This is based on current system settings.
         */
        getNumberDecimalSeparator(): string;

        /**
         * Gets the string used to separate groups of digits to the left of the decimal for numeric values. This is based on current system settings.
         */
        getNumberGroupSeparator(): string;
    }

    /**
     * Defines the culturally appropriate format of displaying numbers. This is based on current system culture settings.
     */
    interface DatetimeFormatInfo {
        /**
         * Gets the string used as the date separator. This is based on current system settings.
         */
        getDateSeparator(): string;

        /**
         * Gets the format string for a long date value. This is based on current system settings.
         */
        getLongDatePattern(): string;

        /**
         * Gets the format string for a long time value. This is based on current system settings.
         */
        getLongTimePattern(): string;

        /**
         * Gets the format string for a short date value. This is based on current system settings.
         */
        getShortDatePattern(): string;

        /**
         * Gets the string used as the time separator. This is based on current system settings.
         */
        getTimeSeparator(): string;
    }

    /**
     * Represents a custom XML part object in a workbook.
     */
    interface CustomXmlPart {
        /**
         * The custom XML part's ID.
         */
        getId(): string;

        /**
         * The custom XML part's namespace URI.
         */
        getNamespaceUri(): string;

        /**
         * Deletes the custom XML part.
         */
        delete(): void;

        /**
         * Gets the custom XML part's full XML content.
         */
        getXml(): string;

        /**
         * Sets the custom XML part's full XML content.
         * @param xml XML content for the part.
         */
        setXml(xml: string): void;
    }

    /**
     * Represents an Excel PivotTable.
     */
    interface PivotTable {
        /**
         * Specifies if the PivotTable allows the application of multiple PivotFilters on a given PivotField in the table.
         */
        getAllowMultipleFiltersPerField(): boolean;

        /**
         * Specifies if the PivotTable allows the application of multiple PivotFilters on a given PivotField in the table.
         */
        setAllowMultipleFiltersPerField(
            allowMultipleFiltersPerField: boolean
        ): void;

        /**
         * Specifies if the PivotTable allows values in the data body to be edited by the user.
         */
        getEnableDataValueEditing(): boolean;

        /**
         * Specifies if the PivotTable allows values in the data body to be edited by the user.
         */
        setEnableDataValueEditing(enableDataValueEditing: boolean): void;

        /**
         * ID of the PivotTable.
         */
        getId(): string;

        /**
         * The PivotLayout describing the layout and visual structure of the PivotTable.
         */
        getLayout(): PivotLayout;

        /**
         * Name of the PivotTable.
         */
        getName(): string;

        /**
         * Name of the PivotTable.
         */
        setName(name: string): void;

        /**
         * Specifies if the PivotTable uses custom lists when sorting.
         */
        getUseCustomSortLists(): boolean;

        /**
         * Specifies if the PivotTable uses custom lists when sorting.
         */
        setUseCustomSortLists(useCustomSortLists: boolean): void;

        /**
         * The worksheet containing the current PivotTable.
         */
        getWorksheet(): Worksheet;

        /**
         * Deletes the PivotTable.
         */
        delete(): void;

        /**
         * Refreshes the PivotTable.
         */
        refresh(): void;

        /**
         * The Column Pivot Hierarchies of the PivotTable.
         */
        getColumnHierarchies(): RowColumnPivotHierarchy[];

        /**
         * Adds the PivotHierarchy to the current axis. If the hierarchy is present elsewhere on the row, column,
         * or filter axis, it will be removed from that location.
         */
        addColumnHierarchy(
            pivotHierarchy: PivotHierarchy
        ): RowColumnPivotHierarchy;

        /**
         * Gets a RowColumnPivotHierarchy by name. If the RowColumnPivotHierarchy does not exist, then this method returns `undefined`.
         * @param name Name of the RowColumnPivotHierarchy to be retrieved.
         */
        getColumnHierarchy(name: string): RowColumnPivotHierarchy | undefined;

        /**
         * Removes the PivotHierarchy from the current axis.
         */
        removeColumnHierarchy(
            rowColumnPivotHierarchy: RowColumnPivotHierarchy
        ): void;

        /**
         * The Data Pivot Hierarchies of the PivotTable.
         */
        getDataHierarchies(): DataPivotHierarchy[];

        /**
         * Adds the PivotHierarchy to the current axis.
         */
        addDataHierarchy(pivotHierarchy: PivotHierarchy): DataPivotHierarchy;

        /**
         * Gets a DataPivotHierarchy by name. If the DataPivotHierarchy does not exist, then this method returns `undefined`.
         * @param name Name of the DataPivotHierarchy to be retrieved.
         */
        getDataHierarchy(name: string): DataPivotHierarchy | undefined;

        /**
         * Removes the PivotHierarchy from the current axis.
         */
        removeDataHierarchy(DataPivotHierarchy: DataPivotHierarchy): void;

        /**
         * The Filter Pivot Hierarchies of the PivotTable.
         */
        getFilterHierarchies(): FilterPivotHierarchy[];

        /**
         * Adds the PivotHierarchy to the current axis. If the hierarchy is present elsewhere on the row, column,
         * or filter axis, it will be removed from that location.
         */
        addFilterHierarchy(
            pivotHierarchy: PivotHierarchy
        ): FilterPivotHierarchy;

        /**
         * Gets a FilterPivotHierarchy by name. If the FilterPivotHierarchy does not exist, then this method returns `undefined`.
         * @param name Name of the FilterPivotHierarchy to be retrieved.
         */
        getFilterHierarchy(name: string): FilterPivotHierarchy | undefined;

        /**
         * Removes the PivotHierarchy from the current axis.
         */
        removeFilterHierarchy(filterPivotHierarchy: FilterPivotHierarchy): void;

        /**
         * The Pivot Hierarchies of the PivotTable.
         */
        getHierarchies(): PivotHierarchy[];

        /**
         * Gets a PivotHierarchy by name. If the PivotHierarchy does not exist, then this method returns `undefined`.
         * @param name Name of the PivotHierarchy to be retrieved.
         */
        getHierarchy(name: string): PivotHierarchy | undefined;

        /**
         * The Row Pivot Hierarchies of the PivotTable.
         */
        getRowHierarchies(): RowColumnPivotHierarchy[];

        /**
         * Adds the PivotHierarchy to the current axis. If the hierarchy is present elsewhere on the row, column,
         * or filter axis, it will be removed from that location.
         */
        addRowHierarchy(
            pivotHierarchy: PivotHierarchy
        ): RowColumnPivotHierarchy;

        /**
         * Gets a RowColumnPivotHierarchy by name. If the RowColumnPivotHierarchy does not exist, then this method returns `undefined`.
         * @param name Name of the RowColumnPivotHierarchy to be retrieved.
         */
        getRowHierarchy(name: string): RowColumnPivotHierarchy | undefined;

        /**
         * Removes the PivotHierarchy from the current axis.
         */
        removeRowHierarchy(
            rowColumnPivotHierarchy: RowColumnPivotHierarchy
        ): void;
    }

    /**
     * Represents the visual layout of the PivotTable.
     */
    interface PivotLayout {
        /**
         * Specifies if formatting will be automatically formatted when it’s refreshed or when fields are moved.
         */
        getAutoFormat(): boolean;

        /**
         * Specifies if formatting will be automatically formatted when it’s refreshed or when fields are moved.
         */
        setAutoFormat(autoFormat: boolean): void;

        /**
         * Specifies if the field list can be shown in the UI.
         */
        getEnableFieldList(): boolean;

        /**
         * Specifies if the field list can be shown in the UI.
         */
        setEnableFieldList(enableFieldList: boolean): void;

        /**
         * This property indicates the PivotLayoutType of all fields on the PivotTable. If fields have different states, this will be null.
         */
        getLayoutType(): PivotLayoutType;

        /**
         * This property indicates the PivotLayoutType of all fields on the PivotTable. If fields have different states, this will be null.
         */
        setLayoutType(layoutType: PivotLayoutType): void;

        /**
         * Specifies if formatting is preserved when the report is refreshed or recalculated by operations such as pivoting, sorting, or changing page field items.
         */
        getPreserveFormatting(): boolean;

        /**
         * Specifies if formatting is preserved when the report is refreshed or recalculated by operations such as pivoting, sorting, or changing page field items.
         */
        setPreserveFormatting(preserveFormatting: boolean): void;

        /**
         * Specifies if the PivotTable report shows grand totals for columns.
         */
        getShowColumnGrandTotals(): boolean;

        /**
         * Specifies if the PivotTable report shows grand totals for columns.
         */
        setShowColumnGrandTotals(showColumnGrandTotals: boolean): void;

        /**
         * Specifies if the PivotTable report shows grand totals for rows.
         */
        getShowRowGrandTotals(): boolean;

        /**
         * Specifies if the PivotTable report shows grand totals for rows.
         */
        setShowRowGrandTotals(showRowGrandTotals: boolean): void;

        /**
         * This property indicates the `SubtotalLocationType` of all fields on the PivotTable. If fields have different states, this will be `null`.
         */
        getSubtotalLocation(): SubtotalLocationType;

        /**
         * This property indicates the `SubtotalLocationType` of all fields on the PivotTable. If fields have different states, this will be `null`.
         */
        setSubtotalLocation(subtotalLocation: SubtotalLocationType): void;

        /**
         * Returns the range where the PivotTable's column labels reside.
         */
        getColumnLabelRange(): Range;

        /**
         * Returns the range where the PivotTable's data values reside.
         */
        getBodyAndTotalRange(): Range;

        /**
         * Gets the DataHierarchy that is used to calculate the value in a specified range within the PivotTable.
         * @param cell A single cell within the PivotTable data body.
         */
        getDataHierarchy(cell: Range | string): DataPivotHierarchy;

        /**
         * Returns the range of the PivotTable's filter area.
         */
        getFilterAxisRange(): Range;

        /**
         * Returns the range the PivotTable exists on, excluding the filter area.
         */
        getRange(): Range;

        /**
         * Returns the range where the PivotTable's row labels reside.
         */
        getRowLabelRange(): Range;

        /**
         * Sets the PivotTable to automatically sort using the specified cell to automatically select all necessary criteria and context. This behaves identically to applying an autosort from the UI.
         * @param cell A single cell to use get the criteria from for applying the autosort.
         * @param sortBy The direction of the sort.
         */
        setAutoSortOnCell(cell: Range | string, sortBy: SortBy): void;
    }

    /**
     * Represents the Excel PivotHierarchy.
     */
    interface PivotHierarchy {
        /**
         * ID of the PivotHierarchy.
         */
        getId(): string;

        /**
         * Name of the PivotHierarchy.
         */
        getName(): string;

        /**
         * Name of the PivotHierarchy.
         */
        setName(name: string): void;

        /**
         * Returns the PivotFields associated with the PivotHierarchy.
         */
        getFields(): PivotField[];

        /**
         * Gets a PivotField by name. If the PivotField does not exist, then this method returns `undefined`.
         * @param name Name of the PivotField to be retrieved.
         */
        getPivotField(name: string): PivotField | undefined;
    }

    /**
     * Represents the Excel RowColumnPivotHierarchy.
     */
    interface RowColumnPivotHierarchy {
        /**
         * ID of the RowColumnPivotHierarchy.
         */
        getId(): string;

        /**
         * Name of the RowColumnPivotHierarchy.
         */
        getName(): string;

        /**
         * Name of the RowColumnPivotHierarchy.
         */
        setName(name: string): void;

        /**
         * Position of the RowColumnPivotHierarchy.
         */
        getPosition(): number;

        /**
         * Position of the RowColumnPivotHierarchy.
         */
        setPosition(position: number): void;

        /**
         * Reset the RowColumnPivotHierarchy back to its default values.
         */
        setToDefault(): void;

        /**
         * Returns the PivotFields associated with the RowColumnPivotHierarchy.
         */
        getFields(): PivotField[];

        /**
         * Gets a PivotField by name. If the PivotField does not exist, then this method returns `undefined`.
         * @param name Name of the PivotField to be retrieved.
         */
        getPivotField(name: string): PivotField | undefined;
    }

    /**
     * Represents the Excel FilterPivotHierarchy.
     */
    interface FilterPivotHierarchy {
        /**
         * Determines whether to allow multiple filter items.
         */
        getEnableMultipleFilterItems(): boolean;

        /**
         * Determines whether to allow multiple filter items.
         */
        setEnableMultipleFilterItems(enableMultipleFilterItems: boolean): void;

        /**
         * ID of the FilterPivotHierarchy.
         */
        getId(): string;

        /**
         * Name of the FilterPivotHierarchy.
         */
        getName(): string;

        /**
         * Name of the FilterPivotHierarchy.
         */
        setName(name: string): void;

        /**
         * Position of the FilterPivotHierarchy.
         */
        getPosition(): number;

        /**
         * Position of the FilterPivotHierarchy.
         */
        setPosition(position: number): void;

        /**
         * Reset the FilterPivotHierarchy back to its default values.
         */
        setToDefault(): void;

        /**
         * Returns the PivotFields associated with the FilterPivotHierarchy.
         */
        getFields(): PivotField[];

        /**
         * Gets a PivotField by name. If the PivotField does not exist, then this method returns `undefined`.
         * @param name Name of the PivotField to be retrieved.
         */
        getPivotField(name: string): PivotField | undefined;
    }

    /**
     * Represents the Excel DataPivotHierarchy.
     */
    interface DataPivotHierarchy {
        /**
         * Returns the PivotFields associated with the DataPivotHierarchy.
         */
        getField(): PivotField;

        /**
         * ID of the DataPivotHierarchy.
         */
        getId(): string;

        /**
         * Name of the DataPivotHierarchy.
         */
        getName(): string;

        /**
         * Name of the DataPivotHierarchy.
         */
        setName(name: string): void;

        /**
         * Number format of the DataPivotHierarchy.
         */
        getNumberFormat(): string;

        /**
         * Number format of the DataPivotHierarchy.
         */
        setNumberFormat(numberFormat: string): void;

        /**
         * Position of the DataPivotHierarchy.
         */
        getPosition(): number;

        /**
         * Position of the DataPivotHierarchy.
         */
        setPosition(position: number): void;

        /**
         * Specifies if the data should be shown as a specific summary calculation.
         */
        getShowAs(): ShowAsRule;

        /**
         * Specifies if the data should be shown as a specific summary calculation.
         */
        setShowAs(showAs: ShowAsRule): void;

        /**
         * Specifies if all items of the DataPivotHierarchy are shown.
         */
        getSummarizeBy(): AggregationFunction;

        /**
         * Specifies if all items of the DataPivotHierarchy are shown.
         */
        setSummarizeBy(summarizeBy: AggregationFunction): void;

        /**
         * Reset the DataPivotHierarchy back to its default values.
         */
        setToDefault(): void;
    }

    /**
     * Represents the Excel PivotField.
     */
    interface PivotField {
        /**
         * ID of the PivotField.
         */
        getId(): string;

        /**
         * Name of the PivotField.
         */
        getName(): string;

        /**
         * Name of the PivotField.
         */
        setName(name: string): void;

        /**
         * Determines whether to show all items of the PivotField.
         */
        getShowAllItems(): boolean;

        /**
         * Determines whether to show all items of the PivotField.
         */
        setShowAllItems(showAllItems: boolean): void;

        /**
         * Subtotals of the PivotField.
         */
        getSubtotals(): Subtotals;

        /**
         * Subtotals of the PivotField.
         */
        setSubtotals(subtotals: Subtotals): void;

        /**
         * Sets one or more of the field's current PivotFilters and applies them to the field.
         * If the provided filters are invalid or cannot be applied, an exception is thrown.
         * @param filter A configured specific PivotFilter, or a PivotFilters interface containing multiple configured filters.
         */
        applyFilter(filter: PivotFilters): void;

        /**
         * Clears all criteria from all of the field's filters. This removes any active filtering on the field.
         */
        clearAllFilters(): void;

        /**
         * Clears all existing criteria from the field's filter of the given type (if one is currently applied).
         * @param filterType The type of filter on the field of which to clear all criteria.
         */
        clearFilter(filterType: PivotFilterType): void;

        /**
         * Gets all filters currently applied on the field.
         */
        getFilters(): PivotFilters;

        /**
         * Checks if there are any applied filters on the field.
         * @param filterType The filter type to check. If no type is provided, this method will check if any filter is applied.
         */
        isFiltered(filterType?: PivotFilterType): boolean;

        /**
         * Sorts the PivotField. If a DataPivotHierarchy is specified, then sort will be applied based on it, if not sort will be based on the PivotField itself.
         * @param sortBy Specifies if the sorting is done in ascending or descending order.
         */
        sortByLabels(sortBy: SortBy): void;

        /**
         * Sorts the PivotField by specified values in a given scope. The scope defines which specific values will be used to sort when
         * there are multiple values from the same DataPivotHierarchy.
         * @param sortBy Specifies if the sorting is done in ascending or descending order.
         * @param valuesHierarchy Specifies the values hierarchy on the data axis to be used for sorting.
         * @param pivotItemScope The items that should be used for the scope of the sorting. These will be the
         * items that make up the row or column that you want to sort on. If a string is used instead of a PivotItem,
         * the string represents the ID of the PivotItem. If there are no items other than data hierarchy on the axis
         * you want to sort on, this can be empty.
         */
        sortByValues(
            sortBy: SortBy,
            valuesHierarchy: DataPivotHierarchy,
            pivotItemScope?: Array<PivotItem | string>
        ): void;

        /**
         * Returns the PivotItems associated with the PivotField.
         */
        getItems(): PivotItem[];

        /**
         * Gets a PivotItem by name. If the PivotItem does not exist, then this method returns `undefined`.
         * @param name Name of the PivotItem to be retrieved.
         */
        getPivotItem(name: string): PivotItem | undefined;
    }

    /**
     * Represents the Excel PivotItem.
     */
    interface PivotItem {
        /**
         * ID of the PivotItem.
         */
        getId(): string;

        /**
         * Determines whether the item is expanded to show child items or if it's collapsed and child items are hidden.
         */
        getIsExpanded(): boolean;

        /**
         * Determines whether the item is expanded to show child items or if it's collapsed and child items are hidden.
         */
        setIsExpanded(isExpanded: boolean): void;

        /**
         * Name of the PivotItem.
         */
        getName(): string;

        /**
         * Name of the PivotItem.
         */
        setName(name: string): void;

        /**
         * Specifies if the PivotItem is visible.
         */
        getVisible(): boolean;

        /**
         * Specifies if the PivotItem is visible.
         */
        setVisible(visible: boolean): void;
    }

    /**
     * Represents a worksheet-level custom property.
     */
    interface WorksheetCustomProperty {
        /**
         * Gets the key of the custom property. Custom property keys are case-insensitive. The key is limited to 255 characters (larger values will cause an `InvalidArgument` error to be thrown.)
         */
        getKey(): string;

        /**
         * Gets or sets the value of the custom property.
         */
        getValue(): string;

        /**
         * Gets or sets the value of the custom property.
         */
        setValue(value: string): void;

        /**
         * Deletes the custom property.
         */
        delete(): void;
    }

    /**
     * Represents workbook properties.
     */
    interface DocumentProperties {
        /**
         * The author of the workbook.
         */
        getAuthor(): string;

        /**
         * The author of the workbook.
         */
        setAuthor(author: string): void;

        /**
         * The category of the workbook.
         */
        getCategory(): string;

        /**
         * The category of the workbook.
         */
        setCategory(category: string): void;

        /**
         * The comments of the workbook.
         */
        getComments(): string;

        /**
         * The comments of the workbook.
         */
        setComments(comments: string): void;

        /**
         * The company of the workbook.
         */
        getCompany(): string;

        /**
         * The company of the workbook.
         */
        setCompany(company: string): void;

        /**
         * Gets the creation date of the workbook.
         */
        getCreationDate(): Date;

        /**
         * The keywords of the workbook.
         */
        getKeywords(): string;

        /**
         * The keywords of the workbook.
         */
        setKeywords(keywords: string): void;

        /**
         * Gets the last author of the workbook.
         */
        getLastAuthor(): string;

        /**
         * The manager of the workbook.
         */
        getManager(): string;

        /**
         * The manager of the workbook.
         */
        setManager(manager: string): void;

        /**
         * Gets the revision number of the workbook.
         */
        getRevisionNumber(): number;

        /**
         * Gets the revision number of the workbook.
         */
        setRevisionNumber(revisionNumber: number): void;

        /**
         * The subject of the workbook.
         */
        getSubject(): string;

        /**
         * The subject of the workbook.
         */
        setSubject(subject: string): void;

        /**
         * The title of the workbook.
         */
        getTitle(): string;

        /**
         * The title of the workbook.
         */
        setTitle(title: string): void;

        /**
         * Gets the collection of custom properties of the workbook.
         */
        getCustom(): CustomProperty[];

        /**
         * Creates a new or sets an existing custom property.
         * @param key Required. The custom property's key, which is case-insensitive. The key is limited to 255 characters outside of Excel on the web (larger keys are automatically trimmed to 255 characters on other platforms).
         * @param value Required. The custom property's value. The value is limited to 255 characters outside of Excel on the web (larger values are automatically trimmed to 255 characters on other platforms).
         */
        addCustomProperty(key: string, value: any): CustomProperty;

        /**
         * Deletes all custom properties in this collection.
         */
        deleteAllCustomProperties(): void;

        /**
         * Gets a custom property object by its key, which is case-insensitive. If the custom property doesn't exist, then this method returns `undefined`.
         * @param key Required. The key that identifies the custom property object.
         */
        getCustomProperty(key: string): CustomProperty | undefined;
    }

    /**
     * Represents a custom property.
     */
    interface CustomProperty {
        /**
         * The key of the custom property. The key is limited to 255 characters outside of Excel on the web (larger keys are automatically trimmed to 255 characters on other platforms).
         */
        getKey(): string;

        /**
         * The type of the value used for the custom property.
         */
        getType(): DocumentPropertyType;

        /**
         * The value of the custom property. The value is limited to 255 characters outside of Excel on the web (larger values are automatically trimmed to 255 characters on other platforms).
         */
        getValue(): any;

        /**
         * The value of the custom property. The value is limited to 255 characters outside of Excel on the web (larger values are automatically trimmed to 255 characters on other platforms).
         */
        setValue(value: any): void;

        /**
         * Deletes the custom property.
         */
        delete(): void;
    }

    /**
     * An object encapsulating a conditional format's range, format, rule, and other properties.
     */
    interface ConditionalFormat {
        /**
         * Returns the cell value conditional format properties if the current conditional format is a `CellValue` type.
         */
        getCellValue(): CellValueConditionalFormat | undefined;

        /**
         * Returns the color scale conditional format properties if the current conditional format is a `ColorScale` type.
         */
        getColorScale(): ColorScaleConditionalFormat | undefined;

        /**
         * Returns the custom conditional format properties if the current conditional format is a custom type.
         */
        getCustom(): CustomConditionalFormat | undefined;

        /**
         * Returns the data bar properties if the current conditional format is a data bar.
         */
        getDataBar(): DataBarConditionalFormat | undefined;

        /**
         * Returns the icon set conditional format properties if the current conditional format is an `IconSet` type.
         */
        getIconSet(): IconSetConditionalFormat | undefined;

        /**
         * The priority of the conditional format in the current `ConditionalFormatCollection`.
         */
        getId(): string;

        /**
         * Returns the preset criteria conditional format. See `ExcelScript.PresetCriteriaConditionalFormat` for more details.
         */
        getPreset(): PresetCriteriaConditionalFormat | undefined;

        /**
         * The priority (or index) within the conditional format collection that this conditional format currently exists in. Changing this also
         * changes other conditional formats' priorities, to allow for a contiguous priority order.
         * Use a negative priority to begin from the back.
         * Priorities greater than the bounds will get and set to the maximum (or minimum if negative) priority.
         * Also note that if you change the priority, you have to re-fetch a new copy of the object at that new priority location if you want to make further changes to it.
         */
        getPriority(): number;

        /**
         * The priority (or index) within the conditional format collection that this conditional format currently exists in. Changing this also
         * changes other conditional formats' priorities, to allow for a contiguous priority order.
         * Use a negative priority to begin from the back.
         * Priorities greater than the bounds will get and set to the maximum (or minimum if negative) priority.
         * Also note that if you change the priority, you have to re-fetch a new copy of the object at that new priority location if you want to make further changes to it.
         */
        setPriority(priority: number): void;

        /**
         * If the conditions of this conditional format are met, no lower-priority formats shall take effect on that cell.
         * Value is `null` on data bars, icon sets, and color scales as there's no concept of `StopIfTrue` for these.
         */
        getStopIfTrue(): boolean;

        /**
         * If the conditions of this conditional format are met, no lower-priority formats shall take effect on that cell.
         * Value is `null` on data bars, icon sets, and color scales as there's no concept of `StopIfTrue` for these.
         */
        setStopIfTrue(stopIfTrue: boolean): void;

        /**
         * Returns the specific text conditional format properties if the current conditional format is a text type.
         * For example, to format cells matching the word "Text".
         */
        getTextComparison(): TextConditionalFormat | undefined;

        /**
         * Returns the top/bottom conditional format properties if the current conditional format is a `TopBottom` type.
         * For example, to format the top 10% or bottom 10 items.
         */
        getTopBottom(): TopBottomConditionalFormat | undefined;

        /**
         * A type of conditional format. Only one can be set at a time.
         */
        getType(): ConditionalFormatType;

        /**
         * Deletes this conditional format.
         */
        delete(): void;

        /**
         * Returns the range to which the conditional format is applied. If the conditional format is applied to multiple ranges, then this method returns `undefined`.
         */
        getRange(): Range;

        /**
         * Returns the `RangeAreas`, comprising one or more rectangular ranges, to which the conditional format is applied.
         */
        getRanges(): RangeAreas;
    }

    /**
     * Represents an Excel conditional data bar type.
     */
    interface DataBarConditionalFormat {
        /**
         * HTML color code representing the color of the Axis line, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
         * Value is "" (an empty string) if no axis is present or set.
         */
        getAxisColor(): string;

        /**
         * HTML color code representing the color of the Axis line, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
         * Value is "" (an empty string) if no axis is present or set.
         */
        setAxisColor(axisColor: string): void;

        /**
         * Representation of how the axis is determined for an Excel data bar.
         */
        getAxisFormat(): ConditionalDataBarAxisFormat;

        /**
         * Representation of how the axis is determined for an Excel data bar.
         */
        setAxisFormat(axisFormat: ConditionalDataBarAxisFormat): void;

        /**
         * Specifies the direction that the data bar graphic should be based on.
         */
        getBarDirection(): ConditionalDataBarDirection;

        /**
         * Specifies the direction that the data bar graphic should be based on.
         */
        setBarDirection(barDirection: ConditionalDataBarDirection): void;

        /**
         * The rule for what consistutes the lower bound (and how to calculate it, if applicable) for a data bar.
         * The `ConditionalDataBarRule` object must be set as a JSON object (use `x.lowerBoundRule = {...}` instead of `x.lowerBoundRule.formula = ...`).
         */
        getLowerBoundRule(): ConditionalDataBarRule;

        /**
         * The rule for what consistutes the lower bound (and how to calculate it, if applicable) for a data bar.
         * The `ConditionalDataBarRule` object must be set as a JSON object (use `x.lowerBoundRule = {...}` instead of `x.lowerBoundRule.formula = ...`).
         */
        setLowerBoundRule(lowerBoundRule: ConditionalDataBarRule): void;

        /**
         * Representation of all values to the left of the axis in an Excel data bar.
         */
        getNegativeFormat(): ConditionalDataBarNegativeFormat;

        /**
         * Representation of all values to the right of the axis in an Excel data bar.
         */
        getPositiveFormat(): ConditionalDataBarPositiveFormat;

        /**
         * If `true`, hides the values from the cells where the data bar is applied.
         */
        getShowDataBarOnly(): boolean;

        /**
         * If `true`, hides the values from the cells where the data bar is applied.
         */
        setShowDataBarOnly(showDataBarOnly: boolean): void;

        /**
         * The rule for what constitutes the upper bound (and how to calculate it, if applicable) for a data bar.
         * The `ConditionalDataBarRule` object must be set as a JSON object (use `x.upperBoundRule = {...}` instead of `x.upperBoundRule.formula = ...`).
         */
        getUpperBoundRule(): ConditionalDataBarRule;

        /**
         * The rule for what constitutes the upper bound (and how to calculate it, if applicable) for a data bar.
         * The `ConditionalDataBarRule` object must be set as a JSON object (use `x.upperBoundRule = {...}` instead of `x.upperBoundRule.formula = ...`).
         */
        setUpperBoundRule(upperBoundRule: ConditionalDataBarRule): void;
    }

    /**
     * Represents a conditional format for the positive side of the data bar.
     */
    interface ConditionalDataBarPositiveFormat {
        /**
         * HTML color code representing the color of the border line, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
         * Value is "" (an empty string) if no border is present or set.
         */
        getBorderColor(): string;

        /**
         * HTML color code representing the color of the border line, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
         * Value is "" (an empty string) if no border is present or set.
         */
        setBorderColor(borderColor: string): void;

        /**
         * HTML color code representing the fill color, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
         */
        getFillColor(): string;

        /**
         * HTML color code representing the fill color, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
         */
        setFillColor(fillColor: string): void;

        /**
         * Specifies if the data bar has a gradient.
         */
        getGradientFill(): boolean;

        /**
         * Specifies if the data bar has a gradient.
         */
        setGradientFill(gradientFill: boolean): void;
    }

    /**
     * Represents a conditional format for the negative side of the data bar.
     */
    interface ConditionalDataBarNegativeFormat {
        /**
         * HTML color code representing the color of the border line, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
         * Value is "" (an empty string) if no border is present or set.
         */
        getBorderColor(): string;

        /**
         * HTML color code representing the color of the border line, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
         * Value is "" (an empty string) if no border is present or set.
         */
        setBorderColor(borderColor: string): void;

        /**
         * HTML color code representing the fill color, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
         */
        getFillColor(): string;

        /**
         * HTML color code representing the fill color, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
         */
        setFillColor(fillColor: string): void;

        /**
         * Specifies if the negative data bar has the same border color as the positive data bar.
         */
        getMatchPositiveBorderColor(): boolean;

        /**
         * Specifies if the negative data bar has the same border color as the positive data bar.
         */
        setMatchPositiveBorderColor(matchPositiveBorderColor: boolean): void;

        /**
         * Specifies if the negative data bar has the same fill color as the positive data bar.
         */
        getMatchPositiveFillColor(): boolean;

        /**
         * Specifies if the negative data bar has the same fill color as the positive data bar.
         */
        setMatchPositiveFillColor(matchPositiveFillColor: boolean): void;
    }

    /**
     * Represents a custom conditional format type.
     */
    interface CustomConditionalFormat {
        /**
         * Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties.
         */
        getFormat(): ConditionalRangeFormat;

        /**
         * Specifies the `Rule` object on this conditional format.
         */
        getRule(): ConditionalFormatRule;
    }

    /**
     * Represents a rule, for all traditional rule/format pairings.
     */
    interface ConditionalFormatRule {
        /**
         * The formula, if required, on which to evaluate the conditional format rule.
         */
        getFormula(): string;

        /**
         * The formula, if required, on which to evaluate the conditional format rule.
         */
        setFormula(formula: string): void;

        /**
         * The formula, if required, on which to evaluate the conditional format rule in the user's language.
         */
        getFormulaLocal(): string;

        /**
         * The formula, if required, on which to evaluate the conditional format rule in the user's language.
         */
        setFormulaLocal(formulaLocal: string): void;
    }

    /**
     * Represents an icon set criteria for conditional formatting.
     */
    interface IconSetConditionalFormat {
        /**
         * An array of criteria and icon sets for the rules and potential custom icons for conditional icons. Note that for the first criterion only the custom icon can be modified, while type, formula, and operator will be ignored when set.
         */
        getCriteria(): ConditionalIconCriterion[];

        /**
         * An array of criteria and icon sets for the rules and potential custom icons for conditional icons. Note that for the first criterion only the custom icon can be modified, while type, formula, and operator will be ignored when set.
         */
        setCriteria(criteria: ConditionalIconCriterion[]): void;

        /**
         * If `true`, reverses the icon orders for the icon set. Note that this cannot be set if custom icons are used.
         */
        getReverseIconOrder(): boolean;

        /**
         * If `true`, reverses the icon orders for the icon set. Note that this cannot be set if custom icons are used.
         */
        setReverseIconOrder(reverseIconOrder: boolean): void;

        /**
         * If `true`, hides the values and only shows icons.
         */
        getShowIconOnly(): boolean;

        /**
         * If `true`, hides the values and only shows icons.
         */
        setShowIconOnly(showIconOnly: boolean): void;

        /**
         * If set, displays the icon set option for the conditional format.
         */
        getStyle(): IconSet;

        /**
         * If set, displays the icon set option for the conditional format.
         */
        setStyle(style: IconSet): void;
    }

    /**
     * Represents the color scale criteria for conditional formatting.
     */
    interface ColorScaleConditionalFormat {
        /**
         * The criteria of the color scale. Midpoint is optional when using a two point color scale.
         */
        getCriteria(): ConditionalColorScaleCriteria;

        /**
         * The criteria of the color scale. Midpoint is optional when using a two point color scale.
         */
        setCriteria(criteria: ConditionalColorScaleCriteria): void;

        /**
         * If `true`, the color scale will have three points (minimum, midpoint, maximum), otherwise it will have two (minimum, maximum).
         */
        getThreeColorScale(): boolean;
    }

    /**
     * Represents a top/bottom conditional format.
     */
    interface TopBottomConditionalFormat {
        /**
         * Returns a format object, encapsulating the conditional format's font, fill, borders, and other properties.
         */
        getFormat(): ConditionalRangeFormat;

        /**
         * The criteria of the top/bottom conditional format.
         */
        getRule(): ConditionalTopBottomRule;

        /**
         * The criteria of the top/bottom conditional format.
         */
        setRule(rule: ConditionalTopBottomRule): void;
    }

    /**
     * Represents the the preset criteria conditional format such as above average, below average, unique values, contains blank, nonblank, error, and noerror.
     */
    interface PresetCriteriaConditionalFormat {
        /**
         * Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties.
         */
        getFormat(): ConditionalRangeFormat;

        /**
         * The rule of the conditional format.
         */
        getRule(): ConditionalPresetCriteriaRule;

        /**
         * The rule of the conditional format.
         */
        setRule(rule: ConditionalPresetCriteriaRule): void;
    }

    /**
     * Represents a specific text conditional format.
     */
    interface TextConditionalFormat {
        /**
         * Returns a format object, encapsulating the conditional format's font, fill, borders, and other properties.
         */
        getFormat(): ConditionalRangeFormat;

        /**
         * The rule of the conditional format.
         */
        getRule(): ConditionalTextComparisonRule;

        /**
         * The rule of the conditional format.
         */
        setRule(rule: ConditionalTextComparisonRule): void;
    }

    /**
     * Represents a cell value conditional format.
     */
    interface CellValueConditionalFormat {
        /**
         * Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties.
         */
        getFormat(): ConditionalRangeFormat;

        /**
         * Specifies the rule object on this conditional format.
         */
        getRule(): ConditionalCellValueRule;

        /**
         * Specifies the rule object on this conditional format.
         */
        setRule(rule: ConditionalCellValueRule): void;
    }

    /**
     * A format object encapsulating the conditional formats range's font, fill, borders, and other properties.
     */
    interface ConditionalRangeFormat {
        /**
         * Returns the fill object defined on the overall conditional format range.
         */
        getFill(): ConditionalRangeFill;

        /**
         * Returns the font object defined on the overall conditional format range.
         */
        getFont(): ConditionalRangeFont;

        /**
         * Represents Excel's number format code for the given range. Cleared if `null` is passed in.
         */
        getNumberFormat(): string;

        /**
         * Represents Excel's number format code for the given range. Cleared if `null` is passed in.
         */
        setNumberFormat(numberFormat: string): void;

        /**
         * Collection of border objects that apply to the overall conditional format range.
         */
        getBorders(): ConditionalRangeBorder[];

        /**
         * Gets the bottom border.
         */
        getConditionalRangeBorderBottom(): ConditionalRangeBorder;

        /**
         * Gets the left border.
         */
        getConditionalRangeBorderLeft(): ConditionalRangeBorder;

        /**
         * Gets the right border.
         */
        getConditionalRangeBorderRight(): ConditionalRangeBorder;

        /**
         * Gets the top border.
         */
        getConditionalRangeBorderTop(): ConditionalRangeBorder;

        /**
         * Gets a border object using its name.
         * @param index Index value of the border object to be retrieved. See `ExcelScript.ConditionalRangeBorderIndex` for details.
         */
        getConditionalRangeBorder(
            index: ConditionalRangeBorderIndex
        ): ConditionalRangeBorder;
    }

    /**
     * This object represents the font attributes (font style, color, etc.) for an object.
     */
    interface ConditionalRangeFont {
        /**
         * Specifies if the font is bold.
         */
        getBold(): boolean;

        /**
         * Specifies if the font is bold.
         */
        setBold(bold: boolean): void;

        /**
         * HTML color code representation of the text color (e.g., #FF0000 represents Red).
         */
        getColor(): string;

        /**
         * HTML color code representation of the text color (e.g., #FF0000 represents Red).
         */
        setColor(color: string): void;

        /**
         * Specifies if the font is italic.
         */
        getItalic(): boolean;

        /**
         * Specifies if the font is italic.
         */
        setItalic(italic: boolean): void;

        /**
         * Specifies the strikethrough status of the font.
         */
        getStrikethrough(): boolean;

        /**
         * Specifies the strikethrough status of the font.
         */
        setStrikethrough(strikethrough: boolean): void;

        /**
         * The type of underline applied to the font. See `ExcelScript.ConditionalRangeFontUnderlineStyle` for details.
         */
        getUnderline(): ConditionalRangeFontUnderlineStyle;

        /**
         * The type of underline applied to the font. See `ExcelScript.ConditionalRangeFontUnderlineStyle` for details.
         */
        setUnderline(underline: ConditionalRangeFontUnderlineStyle): void;

        /**
         * Resets the font formats.
         */
        clear(): void;
    }

    /**
     * Represents the background of a conditional range object.
     */
    interface ConditionalRangeFill {
        /**
         * HTML color code representing the color of the fill, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
         */
        getColor(): string;

        /**
         * HTML color code representing the color of the fill, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
         */
        setColor(color: string): void;

        /**
         * Resets the fill.
         */
        clear(): void;
    }

    /**
     * Represents the border of an object.
     */
    interface ConditionalRangeBorder {
        /**
         * HTML color code representing the color of the border line, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
         */
        getColor(): string;

        /**
         * HTML color code representing the color of the border line, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
         */
        setColor(color: string): void;

        /**
         * Constant value that indicates the specific side of the border. See `ExcelScript.ConditionalRangeBorderIndex` for details.
         */
        getSideIndex(): ConditionalRangeBorderIndex;

        /**
         * One of the constants of line style specifying the line style for the border. See `ExcelScript.BorderLineStyle` for details.
         */
        getStyle(): ConditionalRangeBorderLineStyle;

        /**
         * One of the constants of line style specifying the line style for the border. See `ExcelScript.BorderLineStyle` for details.
         */
        setStyle(style: ConditionalRangeBorderLineStyle): void;
    }

    /**
     * An object encapsulating a style's format and other properties.
     */
    interface PredefinedCellStyle {
        /**
         * Specifies if text is automatically indented when the text alignment in a cell is set to equal distribution.
         */
        getAutoIndent(): boolean;

        /**
         * Specifies if text is automatically indented when the text alignment in a cell is set to equal distribution.
         */
        setAutoIndent(autoIndent: boolean): void;

        /**
         * Specifies if the style is a built-in style.
         */
        getBuiltIn(): boolean;

        /**
         * The fill of the style.
         */
        getFill(): RangeFill;

        /**
         * A `Font` object that represents the font of the style.
         */
        getFont(): RangeFont;

        /**
         * Specifies if the formula will be hidden when the worksheet is protected.
         */
        getFormulaHidden(): boolean;

        /**
         * Specifies if the formula will be hidden when the worksheet is protected.
         */
        setFormulaHidden(formulaHidden: boolean): void;

        /**
         * Represents the horizontal alignment for the style. See `ExcelScript.HorizontalAlignment` for details.
         */
        getHorizontalAlignment(): HorizontalAlignment;

        /**
         * Represents the horizontal alignment for the style. See `ExcelScript.HorizontalAlignment` for details.
         */
        setHorizontalAlignment(horizontalAlignment: HorizontalAlignment): void;

        /**
         * Specifies if the style includes the auto indent, horizontal alignment, vertical alignment, wrap text, indent level, and text orientation properties.
         */
        getIncludeAlignment(): boolean;

        /**
         * Specifies if the style includes the auto indent, horizontal alignment, vertical alignment, wrap text, indent level, and text orientation properties.
         */
        setIncludeAlignment(includeAlignment: boolean): void;

        /**
         * Specifies if the style includes the color, color index, line style, and weight border properties.
         */
        getIncludeBorder(): boolean;

        /**
         * Specifies if the style includes the color, color index, line style, and weight border properties.
         */
        setIncludeBorder(includeBorder: boolean): void;

        /**
         * Specifies if the style includes the background, bold, color, color index, font style, italic, name, size, strikethrough, subscript, superscript, and underline font properties.
         */
        getIncludeFont(): boolean;

        /**
         * Specifies if the style includes the background, bold, color, color index, font style, italic, name, size, strikethrough, subscript, superscript, and underline font properties.
         */
        setIncludeFont(includeFont: boolean): void;

        /**
         * Specifies if the style includes the number format property.
         */
        getIncludeNumber(): boolean;

        /**
         * Specifies if the style includes the number format property.
         */
        setIncludeNumber(includeNumber: boolean): void;

        /**
         * Specifies if the style includes the color, color index, invert if negative, pattern, pattern color, and pattern color index interior properties.
         */
        getIncludePatterns(): boolean;

        /**
         * Specifies if the style includes the color, color index, invert if negative, pattern, pattern color, and pattern color index interior properties.
         */
        setIncludePatterns(includePatterns: boolean): void;

        /**
         * Specifies if the style includes the formula hidden and locked protection properties.
         */
        getIncludeProtection(): boolean;

        /**
         * Specifies if the style includes the formula hidden and locked protection properties.
         */
        setIncludeProtection(includeProtection: boolean): void;

        /**
         * An integer from 0 to 250 that indicates the indent level for the style.
         */
        getIndentLevel(): number;

        /**
         * An integer from 0 to 250 that indicates the indent level for the style.
         */
        setIndentLevel(indentLevel: number): void;

        /**
         * Specifies if the object is locked when the worksheet is protected.
         */
        getLocked(): boolean;

        /**
         * Specifies if the object is locked when the worksheet is protected.
         */
        setLocked(locked: boolean): void;

        /**
         * The name of the style.
         */
        getName(): string;

        /**
         * The format code of the number format for the style.
         */
        getNumberFormat(): string;

        /**
         * The format code of the number format for the style.
         */
        setNumberFormat(numberFormat: string): void;

        /**
         * The localized format code of the number format for the style.
         */
        getNumberFormatLocal(): string;

        /**
         * The localized format code of the number format for the style.
         */
        setNumberFormatLocal(numberFormatLocal: string): void;

        /**
         * The reading order for the style.
         */
        getReadingOrder(): ReadingOrder;

        /**
         * The reading order for the style.
         */
        setReadingOrder(readingOrder: ReadingOrder): void;

        /**
         * Specifies if text automatically shrinks to fit in the available column width.
         */
        getShrinkToFit(): boolean;

        /**
         * Specifies if text automatically shrinks to fit in the available column width.
         */
        setShrinkToFit(shrinkToFit: boolean): void;

        /**
         * The text orientation for the style.
         */
        getTextOrientation(): number;

        /**
         * The text orientation for the style.
         */
        setTextOrientation(textOrientation: number): void;

        /**
         * Specifies the vertical alignment for the style. See `ExcelScript.VerticalAlignment` for details.
         */
        getVerticalAlignment(): VerticalAlignment;

        /**
         * Specifies the vertical alignment for the style. See `ExcelScript.VerticalAlignment` for details.
         */
        setVerticalAlignment(verticalAlignment: VerticalAlignment): void;

        /**
         * Specifies if Excel wraps the text in the object.
         */
        getWrapText(): boolean;

        /**
         * Specifies if Excel wraps the text in the object.
         */
        setWrapText(wrapText: boolean): void;

        /**
         * Deletes this style.
         */
        delete(): void;

        /**
         * A collection of four border objects that represent the style of the four borders.
         */
        getBorders(): RangeBorder[];

        /**
         * Specifies a double that lightens or darkens a color for range borders. The value is between -1 (darkest) and 1 (brightest), with 0 for the original color.
         * A `null` value indicates that the entire border collection doesn't have a uniform `tintAndShade` setting.
         */
        getRangeBorderTintAndShade(): number;

        /**
         * Specifies a double that lightens or darkens a color for range borders. The value is between -1 (darkest) and 1 (brightest), with 0 for the original color.
         * A `null` value indicates that the entire border collection doesn't have a uniform `tintAndShade` setting.
         */
        setRangeBorderTintAndShade(rangeBorderTintAndShade: number): void;

        /**
         * Gets a border object using its name.
         * @param index Index value of the border object to be retrieved. See `ExcelScript.BorderIndex` for details.
         */
        getRangeBorder(index: BorderIndex): RangeBorder;
    }

    /**
     * Represents a table style, which defines the style elements by region of the table.
     */
    interface TableStyle {
        /**
         * Gets the name of the table style.
         */
        getName(): string;

        /**
         * Gets the name of the table style.
         */
        setName(name: string): void;

        /**
         * Specifies if this `TableStyle` object is read-only.
         */
        getReadOnly(): boolean;

        /**
         * Deletes the table style.
         */
        delete(): void;

        /**
         * Creates a duplicate of this table style with copies of all the style elements.
         */
        duplicate(): TableStyle;
    }

    /**
     * Represents a PivotTable style, which defines style elements by PivotTable region.
     */
    interface PivotTableStyle {
        /**
         * Gets the name of the PivotTable style.
         */
        getName(): string;

        /**
         * Gets the name of the PivotTable style.
         */
        setName(name: string): void;

        /**
         * Specifies if this `PivotTableStyle` object is read-only.
         */
        getReadOnly(): boolean;

        /**
         * Deletes the PivotTable style.
         */
        delete(): void;

        /**
         * Creates a duplicate of this PivotTable style with copies of all the style elements.
         */
        duplicate(): PivotTableStyle;
    }

    /**
     * Represents a slicer style, which defines style elements by region of the slicer.
     */
    interface SlicerStyle {
        /**
         * Gets the name of the slicer style.
         */
        getName(): string;

        /**
         * Gets the name of the slicer style.
         */
        setName(name: string): void;

        /**
         * Specifies if this `SlicerStyle` object is read-only.
         */
        getReadOnly(): boolean;

        /**
         * Deletes the slicer style.
         */
        delete(): void;

        /**
         * Creates a duplicate of this slicer style with copies of all the style elements.
         */
        duplicate(): SlicerStyle;
    }

    /**
     * Represents a `TimelineStyle`, which defines style elements by region in the timeline.
     */
    interface TimelineStyle {
        /**
         * Gets the name of the timeline style.
         */
        getName(): string;

        /**
         * Gets the name of the timeline style.
         */
        setName(name: string): void;

        /**
         * Specifies if this `TimelineStyle` object is read-only.
         */
        getReadOnly(): boolean;

        /**
         * Deletes the table style.
         */
        delete(): void;

        /**
         * Creates a duplicate of this timeline style with copies of all the style elements.
         */
        duplicate(): TimelineStyle;
    }

    /**
     * Represents layout and print settings that are not dependent on any printer-specific implementation. These settings include margins, orientation, page numbering, title rows, and print area.
     */
    interface PageLayout {
        /**
         * The worksheet's black and white print option.
         */
        getBlackAndWhite(): boolean;

        /**
         * The worksheet's black and white print option.
         */
        setBlackAndWhite(blackAndWhite: boolean): void;

        /**
         * The worksheet's bottom page margin to use for printing in points.
         */
        getBottomMargin(): number;

        /**
         * The worksheet's bottom page margin to use for printing in points.
         */
        setBottomMargin(bottomMargin: number): void;

        /**
         * The worksheet's center horizontally flag. This flag determines whether the worksheet will be centered horizontally when it's printed.
         */
        getCenterHorizontally(): boolean;

        /**
         * The worksheet's center horizontally flag. This flag determines whether the worksheet will be centered horizontally when it's printed.
         */
        setCenterHorizontally(centerHorizontally: boolean): void;

        /**
         * The worksheet's center vertically flag. This flag determines whether the worksheet will be centered vertically when it's printed.
         */
        getCenterVertically(): boolean;

        /**
         * The worksheet's center vertically flag. This flag determines whether the worksheet will be centered vertically when it's printed.
         */
        setCenterVertically(centerVertically: boolean): void;

        /**
         * The worksheet's draft mode option. If `true`, the sheet will be printed without graphics.
         */
        getDraftMode(): boolean;

        /**
         * The worksheet's draft mode option. If `true`, the sheet will be printed without graphics.
         */
        setDraftMode(draftMode: boolean): void;

        /**
         * The worksheet's first page number to print. A `null` value represents "auto" page numbering.
         */
        getFirstPageNumber(): number | "";

        /**
         * The worksheet's first page number to print. A `null` value represents "auto" page numbering.
         */
        setFirstPageNumber(firstPageNumber: number | ""): void;

        /**
         * The worksheet's footer margin, in points, for use when printing.
         */
        getFooterMargin(): number;

        /**
         * The worksheet's footer margin, in points, for use when printing.
         */
        setFooterMargin(footerMargin: number): void;

        /**
         * The worksheet's header margin, in points, for use when printing.
         */
        getHeaderMargin(): number;

        /**
         * The worksheet's header margin, in points, for use when printing.
         */
        setHeaderMargin(headerMargin: number): void;

        /**
         * Header and footer configuration for the worksheet.
         */
        getHeadersFooters(): HeaderFooterGroup;

        /**
         * The worksheet's left margin, in points, for use when printing.
         */
        getLeftMargin(): number;

        /**
         * The worksheet's left margin, in points, for use when printing.
         */
        setLeftMargin(leftMargin: number): void;

        /**
         * The worksheet's orientation of the page.
         */
        getOrientation(): PageOrientation;

        /**
         * The worksheet's orientation of the page.
         */
        setOrientation(orientation: PageOrientation): void;

        /**
         * The worksheet's paper size of the page.
         */
        getPaperSize(): PaperType;

        /**
         * The worksheet's paper size of the page.
         */
        setPaperSize(paperSize: PaperType): void;

        /**
         * Specifies if the worksheet's comments should be displayed when printing.
         */
        getPrintComments(): PrintComments;

        /**
         * Specifies if the worksheet's comments should be displayed when printing.
         */
        setPrintComments(printComments: PrintComments): void;

        /**
         * The worksheet's print errors option.
         */
        getPrintErrors(): PrintErrorType;

        /**
         * The worksheet's print errors option.
         */
        setPrintErrors(printErrors: PrintErrorType): void;

        /**
         * Specifies if the worksheet's gridlines will be printed.
         */
        getPrintGridlines(): boolean;

        /**
         * Specifies if the worksheet's gridlines will be printed.
         */
        setPrintGridlines(printGridlines: boolean): void;

        /**
         * Specifies if the worksheet's headings will be printed.
         */
        getPrintHeadings(): boolean;

        /**
         * Specifies if the worksheet's headings will be printed.
         */
        setPrintHeadings(printHeadings: boolean): void;

        /**
         * The worksheet's page print order option. This specifies the order to use for processing the page number printed.
         */
        getPrintOrder(): PrintOrder;

        /**
         * The worksheet's page print order option. This specifies the order to use for processing the page number printed.
         */
        setPrintOrder(printOrder: PrintOrder): void;

        /**
         * The worksheet's right margin, in points, for use when printing.
         */
        getRightMargin(): number;

        /**
         * The worksheet's right margin, in points, for use when printing.
         */
        setRightMargin(rightMargin: number): void;

        /**
         * The worksheet's top margin, in points, for use when printing.
         */
        getTopMargin(): number;

        /**
         * The worksheet's top margin, in points, for use when printing.
         */
        setTopMargin(topMargin: number): void;

        /**
         * The worksheet's print zoom options.
         * The `PageLayoutZoomOptions` object must be set as a JSON object (use `x.zoom = {...}` instead of `x.zoom.scale = ...`).
         */
        getZoom(): PageLayoutZoomOptions;

        /**
         * The worksheet's print zoom options.
         * The `PageLayoutZoomOptions` object must be set as a JSON object (use `x.zoom = {...}` instead of `x.zoom.scale = ...`).
         */
        setZoom(zoom: PageLayoutZoomOptions): void;

        /**
         * Gets the `RangeAreas` object, comprising one or more rectangular ranges, that represents the print area for the worksheet. If there is no print area, then this method returns `undefined`.
         */
        getPrintArea(): RangeAreas;

        /**
         * Gets the range object representing the title columns. If not set, then this method returns `undefined`.
         */
        getPrintTitleColumns(): Range;

        /**
         * Gets the range object representing the title rows. If not set, then this method returns `undefined`.
         */
        getPrintTitleRows(): Range;

        /**
         * Sets the worksheet's print area.
         * @param printArea The range or ranges of the content to print.
         */
        setPrintArea(printArea: Range | RangeAreas | string): void;

        /**
         * Sets the worksheet's page margins with units.
         * @param unit Measurement unit for the margins provided.
         * @param marginOptions Margin values to set. Margins not provided remain unchanged.
         */
        setPrintMargins(
            unit: PrintMarginUnit,
            marginOptions: PageLayoutMarginOptions
        ): void;

        /**
         * Sets the columns that contain the cells to be repeated at the left of each page of the worksheet for printing.
         * @param printTitleColumns The columns to be repeated to the left of each page. The range must span the entire column to be valid.
         */
        setPrintTitleColumns(printTitleColumns: Range | string): void;

        /**
         * Sets the rows that contain the cells to be repeated at the top of each page of the worksheet for printing.
         * @param printTitleRows The rows to be repeated at the top of each page. The range must span the entire row to be valid.
         */
        setPrintTitleRows(printTitleRows: Range | string): void;
    }

    interface HeaderFooter {
        /**
         * The center footer of the worksheet.
         * To apply font formatting or insert a variable value, use format codes specified here: https://msdn.microsoft.com/library/bb225426.aspx.
         */
        getCenterFooter(): string;

        /**
         * The center footer of the worksheet.
         * To apply font formatting or insert a variable value, use format codes specified here: https://msdn.microsoft.com/library/bb225426.aspx.
         */
        setCenterFooter(centerFooter: string): void;

        /**
         * The center header of the worksheet.
         * To apply font formatting or insert a variable value, use format codes specified here: https://msdn.microsoft.com/library/bb225426.aspx.
         */
        getCenterHeader(): string;

        /**
         * The center header of the worksheet.
         * To apply font formatting or insert a variable value, use format codes specified here: https://msdn.microsoft.com/library/bb225426.aspx.
         */
        setCenterHeader(centerHeader: string): void;

        /**
         * The left footer of the worksheet.
         * To apply font formatting or insert a variable value, use format codes specified here: https://msdn.microsoft.com/library/bb225426.aspx.
         */
        getLeftFooter(): string;

        /**
         * The left footer of the worksheet.
         * To apply font formatting or insert a variable value, use format codes specified here: https://msdn.microsoft.com/library/bb225426.aspx.
         */
        setLeftFooter(leftFooter: string): void;

        /**
         * The left header of the worksheet.
         * To apply font formatting or insert a variable value, use format codes specified here: https://msdn.microsoft.com/library/bb225426.aspx.
         */
        getLeftHeader(): string;

        /**
         * The left header of the worksheet.
         * To apply font formatting or insert a variable value, use format codes specified here: https://msdn.microsoft.com/library/bb225426.aspx.
         */
        setLeftHeader(leftHeader: string): void;

        /**
         * The right footer of the worksheet.
         * To apply font formatting or insert a variable value, use format codes specified here: https://msdn.microsoft.com/library/bb225426.aspx.
         */
        getRightFooter(): string;

        /**
         * The right footer of the worksheet.
         * To apply font formatting or insert a variable value, use format codes specified here: https://msdn.microsoft.com/library/bb225426.aspx.
         */
        setRightFooter(rightFooter: string): void;

        /**
         * The right header of the worksheet.
         * To apply font formatting or insert a variable value, use format codes specified here: https://msdn.microsoft.com/library/bb225426.aspx.
         */
        getRightHeader(): string;

        /**
         * The right header of the worksheet.
         * To apply font formatting or insert a variable value, use format codes specified here: https://msdn.microsoft.com/library/bb225426.aspx.
         */
        setRightHeader(rightHeader: string): void;
    }

    interface HeaderFooterGroup {
        /**
         * The general header/footer, used for all pages unless even/odd or first page is specified.
         */
        getDefaultForAllPages(): HeaderFooter;

        /**
         * The header/footer to use for even pages, odd header/footer needs to be specified for odd pages.
         */
        getEvenPages(): HeaderFooter;

        /**
         * The first page header/footer, for all other pages general or even/odd is used.
         */
        getFirstPage(): HeaderFooter;

        /**
         * The header/footer to use for odd pages, even header/footer needs to be specified for even pages.
         */
        getOddPages(): HeaderFooter;

        /**
         * The state by which headers/footers are set. See `ExcelScript.HeaderFooterState` for details.
         */
        getState(): HeaderFooterState;

        /**
         * The state by which headers/footers are set. See `ExcelScript.HeaderFooterState` for details.
         */
        setState(state: HeaderFooterState): void;

        /**
         * Gets or sets a flag indicating if headers/footers are aligned with the page margins set in the page layout options for the worksheet.
         */
        getUseSheetMargins(): boolean;

        /**
         * Gets or sets a flag indicating if headers/footers are aligned with the page margins set in the page layout options for the worksheet.
         */
        setUseSheetMargins(useSheetMargins: boolean): void;

        /**
         * Gets or sets a flag indicating if headers/footers should be scaled by the page percentage scale set in the page layout options for the worksheet.
         */
        getUseSheetScale(): boolean;

        /**
         * Gets or sets a flag indicating if headers/footers should be scaled by the page percentage scale set in the page layout options for the worksheet.
         */
        setUseSheetScale(useSheetScale: boolean): void;
    }

    interface PageBreak {
        /**
         * Specifies the column index for the page break.
         */
        getColumnIndex(): number;

        /**
         * Deletes a page break object.
         */
        delete(): void;

        /**
         * Gets the first cell after the page break.
         */
        getCellAfterBreak(): Range;
    }

    /**
     * Represents a comment in the workbook.
     */
    interface Comment {
        /**
         * Gets the email of the comment's author.
         */
        getAuthorEmail(): string;

        /**
         * Gets the name of the comment's author.
         */
        getAuthorName(): string;

        /**
         * The comment's content. The string is plain text.
         */
        getContent(): string;

        /**
         * The comment's content. The string is plain text.
         */
        setContent(content: string): void;

        /**
         * Gets the content type of the comment.
         */
        getContentType(): ContentType;

        /**
         * Gets the creation time of the comment. Returns `null` if the comment was converted from a note, since the comment does not have a creation date.
         */
        getCreationDate(): Date;

        /**
         * Specifies the comment identifier.
         */
        getId(): string;

        /**
         * Gets the entities (e.g., people) that are mentioned in comments.
         */
        getMentions(): CommentMention[];

        /**
         * The comment thread status. A value of `true` means that the comment thread is resolved.
         */
        getResolved(): boolean;

        /**
         * The comment thread status. A value of `true` means that the comment thread is resolved.
         */
        setResolved(resolved: boolean): void;

        /**
         * Gets the rich comment content (e.g., mentions in comments). This string is not meant to be displayed to end-users. Your add-in should only use this to parse rich comment content.
         */
        getRichContent(): string;

        /**
         * Deletes the comment and all the connected replies.
         */
        delete(): void;

        /**
         * Gets the cell where this comment is located.
         */
        getLocation(): Range;

        /**
         * Updates the comment content with a specially formatted string and a list of mentions.
         * @param contentWithMentions The content for the comment. This contains a specially formatted string and a list of mentions that will be parsed into the string when displayed by Excel.
         */
        updateMentions(contentWithMentions: CommentRichContent): void;

        /**
         * Represents a collection of reply objects associated with the comment.
         */
        getReplies(): CommentReply[];

        /**
         * Creates a comment reply for a comment.
         * @param content The comment's content. This can be either a string or a `CommentRichContent` object (e.g., for comments with mentions).
         * @param contentType Optional. The type of content contained within the comment. The default value is enum `ContentType.Plain`.
         */
        addCommentReply(
            content: CommentRichContent | string,
            contentType?: ContentType
        ): CommentReply;

        /**
         * Returns a comment reply identified by its ID.
         * @param commentReplyId The identifier for the comment reply.
         */
        getCommentReply(commentReplyId: string): CommentReply;
    }

    /**
     * Represents a comment reply in the workbook.
     */
    interface CommentReply {
        /**
         * Gets the email of the comment reply's author.
         */
        getAuthorEmail(): string;

        /**
         * Gets the name of the comment reply's author.
         */
        getAuthorName(): string;

        /**
         * The comment reply's content. The string is plain text.
         */
        getContent(): string;

        /**
         * The comment reply's content. The string is plain text.
         */
        setContent(content: string): void;

        /**
         * The content type of the reply.
         */
        getContentType(): ContentType;

        /**
         * Gets the creation time of the comment reply.
         */
        getCreationDate(): Date;

        /**
         * Specifies the comment reply identifier.
         */
        getId(): string;

        /**
         * The entities (e.g., people) that are mentioned in comments.
         */
        getMentions(): CommentMention[];

        /**
         * The comment reply status. A value of `true` means the reply is in the resolved state.
         */
        getResolved(): boolean;

        /**
         * The rich comment content (e.g., mentions in comments). This string is not meant to be displayed to end-users. Your add-in should only use this to parse rich comment content.
         */
        getRichContent(): string;

        /**
         * Deletes the comment reply.
         */
        delete(): void;

        /**
         * Gets the cell where this comment reply is located.
         */
        getLocation(): Range;

        /**
         * Gets the parent comment of this reply.
         */
        getParentComment(): Comment;

        /**
         * Updates the comment content with a specially formatted string and a list of mentions.
         * @param contentWithMentions The content for the comment. This contains a specially formatted string and a list of mentions that will be parsed into the string when displayed by Excel.
         */
        updateMentions(contentWithMentions: CommentRichContent): void;
    }

    /**
     * Represents a generic shape object in the worksheet. A shape could be a geometric shape, a line, a group of shapes, etc.
     */
    interface Shape {
        /**
         * Specifies the alternative description text for a `Shape` object.
         */
        getAltTextDescription(): string;

        /**
         * Specifies the alternative description text for a `Shape` object.
         */
        setAltTextDescription(altTextDescription: string): void;

        /**
         * Specifies the alternative title text for a `Shape` object.
         */
        getAltTextTitle(): string;

        /**
         * Specifies the alternative title text for a `Shape` object.
         */
        setAltTextTitle(altTextTitle: string): void;

        /**
         * Returns the number of connection sites on this shape.
         */
        getConnectionSiteCount(): number;

        /**
         * Returns the fill formatting of this shape.
         */
        getFill(): ShapeFill;

        /**
         * Returns the geometric shape associated with the shape. An error will be thrown if the shape type is not "GeometricShape".
         */
        getGeometricShape(): GeometricShape;

        /**
         * Specifies the geometric shape type of this geometric shape. See `ExcelScript.GeometricShapeType` for details. Returns `null` if the shape type is not "GeometricShape".
         */
        getGeometricShapeType(): GeometricShapeType;

        /**
         * Specifies the geometric shape type of this geometric shape. See `ExcelScript.GeometricShapeType` for details. Returns `null` if the shape type is not "GeometricShape".
         */
        setGeometricShapeType(geometricShapeType: GeometricShapeType): void;

        /**
         * Returns the shape group associated with the shape. An error will be thrown if the shape type is not "GroupShape".
         */
        getGroup(): ShapeGroup;

        /**
         * Specifies the height, in points, of the shape.
         * Throws an `InvalidArgument` exception when set with a negative value or zero as an input.
         */
        getHeight(): number;

        /**
         * Specifies the height, in points, of the shape.
         * Throws an `InvalidArgument` exception when set with a negative value or zero as an input.
         */
        setHeight(height: number): void;

        /**
         * Specifies the shape identifier.
         */
        getId(): string;

        /**
         * Returns the image associated with the shape. An error will be thrown if the shape type is not "Image".
         */
        getImage(): Image;

        /**
         * The distance, in points, from the left side of the shape to the left side of the worksheet.
         * Throws an `InvalidArgument` exception when set with a negative value as an input.
         */
        getLeft(): number;

        /**
         * The distance, in points, from the left side of the shape to the left side of the worksheet.
         * Throws an `InvalidArgument` exception when set with a negative value as an input.
         */
        setLeft(left: number): void;

        /**
         * Specifies the level of the specified shape. For example, a level of 0 means that the shape is not part of any groups, a level of 1 means the shape is part of a top-level group, and a level of 2 means the shape is part of a sub-group of the top level.
         */
        getLevel(): number;

        /**
         * Returns the line associated with the shape. An error will be thrown if the shape type is not "Line".
         */
        getLine(): Line;

        /**
         * Returns the line formatting of this shape.
         */
        getLineFormat(): ShapeLineFormat;

        /**
         * Specifies if the aspect ratio of this shape is locked.
         */
        getLockAspectRatio(): boolean;

        /**
         * Specifies if the aspect ratio of this shape is locked.
         */
        setLockAspectRatio(lockAspectRatio: boolean): void;

        /**
         * Specifies the name of the shape.
         */
        getName(): string;

        /**
         * Specifies the name of the shape.
         */
        setName(name: string): void;

        /**
         * Specifies the parent group of this shape.
         */
        getParentGroup(): Shape;

        /**
         * Represents how the object is attached to the cells below it.
         */
        getPlacement(): Placement;

        /**
         * Represents how the object is attached to the cells below it.
         */
        setPlacement(placement: Placement): void;

        /**
         * Specifies the rotation, in degrees, of the shape.
         */
        getRotation(): number;

        /**
         * Specifies the rotation, in degrees, of the shape.
         */
        setRotation(rotation: number): void;

        /**
         * Returns the text frame object of this shape.
         */
        getTextFrame(): TextFrame;

        /**
         * The distance, in points, from the top edge of the shape to the top edge of the worksheet.
         * Throws an `InvalidArgument` exception when set with a negative value as an input.
         */
        getTop(): number;

        /**
         * The distance, in points, from the top edge of the shape to the top edge of the worksheet.
         * Throws an `InvalidArgument` exception when set with a negative value as an input.
         */
        setTop(top: number): void;

        /**
         * Returns the type of this shape. See `ExcelScript.ShapeType` for details.
         */
        getType(): ShapeType;

        /**
         * Specifies if the shape is visible.
         */
        getVisible(): boolean;

        /**
         * Specifies if the shape is visible.
         */
        setVisible(visible: boolean): void;

        /**
         * Specifies the width, in points, of the shape.
         * Throws an `InvalidArgument` exception when set with a negative value or zero as an input.
         */
        getWidth(): number;

        /**
         * Specifies the width, in points, of the shape.
         * Throws an `InvalidArgument` exception when set with a negative value or zero as an input.
         */
        setWidth(width: number): void;

        /**
         * Returns the position of the specified shape in the z-order, with 0 representing the bottom of the order stack.
         */
        getZOrderPosition(): number;

        /**
         * Copies and pastes a `Shape` object.
         * The pasted shape is copied to the same pixel location as this shape.
         * @param destinationSheet The sheet to which the shape object will be pasted. The default value is the copied shape's worksheet.
         */
        copyTo(destinationSheet?: Worksheet | string): Shape;

        /**
         * Removes the shape from the worksheet.
         */
        delete(): void;

        /**
         * Converts the shape to an image and returns the image as a base64-encoded string. The DPI is 96. The only supported formats are `ExcelScript.PictureFormat.BMP`, `ExcelScript.PictureFormat.PNG`, `ExcelScript.PictureFormat.JPEG`, and `ExcelScript.PictureFormat.GIF`.
         * @param format Specifies the format of the image.
         */
        getImageAsBase64(format: PictureFormat): string;

        /**
         * Moves the shape horizontally by the specified number of points.
         * @param increment The increment, in points, the shape will be horizontally moved. A positive value moves the shape to the right and a negative value moves it to the left. If the sheet is right-to-left oriented, this is reversed: positive values will move the shape to the left and negative values will move it to the right.
         */
        incrementLeft(increment: number): void;

        /**
         * Rotates the shape clockwise around the z-axis by the specified number of degrees.
         * Use the `rotation` property to set the absolute rotation of the shape.
         * @param increment How many degrees the shape will be rotated. A positive value rotates the shape clockwise and a negative value rotates it counterclockwise.
         */
        incrementRotation(increment: number): void;

        /**
         * Moves the shape vertically by the specified number of points.
         * @param increment The increment, in points, the shape will be vertically moved. A positive value moves the shape down and a negative value moves it up.
         */
        incrementTop(increment: number): void;

        /**
         * Scales the height of the shape by a specified factor. For images, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures are always scaled relative to their current height.
         * @param scaleFactor Specifies the ratio between the height of the shape after you resize it and the current or original height.
         * @param scaleType Specifies whether the shape is scaled relative to its original or current size. The original size scaling option only works for images.
         * @param scaleFrom Optional. Specifies which part of the shape retains its position when the shape is scaled. If omitted, it represents the shape's upper left corner retains its position.
         */
        scaleHeight(
            scaleFactor: number,
            scaleType: ShapeScaleType,
            scaleFrom?: ShapeScaleFrom
        ): void;

        /**
         * Scales the width of the shape by a specified factor. For images, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures are always scaled relative to their current width.
         * @param scaleFactor Specifies the ratio between the width of the shape after you resize it and the current or original width.
         * @param scaleType Specifies whether the shape is scaled relative to its original or current size. The original size scaling option only works for images.
         * @param scaleFrom Optional. Specifies which part of the shape retains its position when the shape is scaled. If omitted, it represents the shape's upper left corner retains its position.
         */
        scaleWidth(
            scaleFactor: number,
            scaleType: ShapeScaleType,
            scaleFrom?: ShapeScaleFrom
        ): void;

        /**
         * Moves the specified shape up or down the collection's z-order, which shifts it in front of or behind other shapes.
         * @param position Where to move the shape in the z-order stack relative to the other shapes. See `ExcelScript.ShapeZOrder` for details.
         */
        setZOrder(position: ShapeZOrder): void;

        /**
         * Converts the shape to an image and returns the image as a base64-encoded string. The DPI is 96. The only supported formats are `ExcelScript.PictureFormat.BMP`, `ExcelScript.PictureFormat.PNG`, `ExcelScript.PictureFormat.JPEG`, and `ExcelScript.PictureFormat.GIF`.
         * @param format Specifies the format of the image.
         * @deprecated Use `getImageAsBase64` instead.
         */
        getAsImage(format: PictureFormat): string;
    }

    /**
     * Represents a geometric shape inside a worksheet. A geometric shape can be a rectangle, block arrow, equation symbol, flowchart item, star, banner, callout, or any other basic shape in Excel.
     */
    interface GeometricShape {
        /**
         * Returns the shape identifier.
         */
        getId(): string;
    }

    /**
     * Represents an image in the worksheet. To get the corresponding `Shape` object, use `Image.getShape`.
     */
    interface Image {
        /**
         * Specifies the shape identifier for the image object.
         */
        getId(): string;

        /**
         * Returns the `Shape` object associated with the image.
         */
        getShape(): Shape;

        /**
         * Returns the format of the image.
         */
        getFormat(): PictureFormat;
    }

    /**
     * Represents a shape group inside a worksheet. To get the corresponding `Shape` object, use `ShapeGroup.shape`.
     */
    interface ShapeGroup {
        /**
         * Specifies the shape identifier.
         */
        getId(): string;

        /**
         * Returns the `Shape` object associated with the group.
         */
        getGroupShape(): Shape;

        /**
         * Ungroups any grouped shapes in the specified shape group.
         */
        ungroup(): void;

        /**
         * Returns the collection of `Shape` objects.
         */
        getShapes(): Shape[];

        /**
         * Gets a shape using its name or ID.
         * @param key The name or ID of the shape to be retrieved.
         */
        getShape(key: string): Shape;
    }

    /**
     * Represents a line inside a worksheet. To get the corresponding `Shape` object, use `Line.shape`.
     */
    interface Line {
        /**
         * Represents the length of the arrowhead at the beginning of the specified line.
         */
        getBeginArrowheadLength(): ArrowheadLength;

        /**
         * Represents the length of the arrowhead at the beginning of the specified line.
         */
        setBeginArrowheadLength(beginArrowheadLength: ArrowheadLength): void;

        /**
         * Represents the style of the arrowhead at the beginning of the specified line.
         */
        getBeginArrowheadStyle(): ArrowheadStyle;

        /**
         * Represents the style of the arrowhead at the beginning of the specified line.
         */
        setBeginArrowheadStyle(beginArrowheadStyle: ArrowheadStyle): void;

        /**
         * Represents the width of the arrowhead at the beginning of the specified line.
         */
        getBeginArrowheadWidth(): ArrowheadWidth;

        /**
         * Represents the width of the arrowhead at the beginning of the specified line.
         */
        setBeginArrowheadWidth(beginArrowheadWidth: ArrowheadWidth): void;

        /**
         * Represents the shape to which the beginning of the specified line is attached.
         */
        getBeginConnectedShape(): Shape;

        /**
         * Represents the connection site to which the beginning of a connector is connected. Returns `null` when the beginning of the line is not attached to any shape.
         */
        getBeginConnectedSite(): number;

        /**
         * Represents the length of the arrowhead at the end of the specified line.
         */
        getEndArrowheadLength(): ArrowheadLength;

        /**
         * Represents the length of the arrowhead at the end of the specified line.
         */
        setEndArrowheadLength(endArrowheadLength: ArrowheadLength): void;

        /**
         * Represents the style of the arrowhead at the end of the specified line.
         */
        getEndArrowheadStyle(): ArrowheadStyle;

        /**
         * Represents the style of the arrowhead at the end of the specified line.
         */
        setEndArrowheadStyle(endArrowheadStyle: ArrowheadStyle): void;

        /**
         * Represents the width of the arrowhead at the end of the specified line.
         */
        getEndArrowheadWidth(): ArrowheadWidth;

        /**
         * Represents the width of the arrowhead at the end of the specified line.
         */
        setEndArrowheadWidth(endArrowheadWidth: ArrowheadWidth): void;

        /**
         * Represents the shape to which the end of the specified line is attached.
         */
        getEndConnectedShape(): Shape;

        /**
         * Represents the connection site to which the end of a connector is connected. Returns `null` when the end of the line is not attached to any shape.
         */
        getEndConnectedSite(): number;

        /**
         * Specifies the shape identifier.
         */
        getId(): string;

        /**
         * Specifies if the beginning of the specified line is connected to a shape.
         */
        getIsBeginConnected(): boolean;

        /**
         * Specifies if the end of the specified line is connected to a shape.
         */
        getIsEndConnected(): boolean;

        /**
         * Returns the `Shape` object associated with the line.
         */
        getShape(): Shape;

        /**
         * Represents the connector type for the line.
         */
        getConnectorType(): ConnectorType;

        /**
         * Represents the connector type for the line.
         */
        setConnectorType(connectorType: ConnectorType): void;

        /**
         * Attaches the beginning of the specified connector to a specified shape.
         * @param shape The shape to connect.
         * @param connectionSite The connection site on the shape to which the beginning of the connector is attached. Must be an integer between 0 (inclusive) and the connection-site count of the specified shape (exclusive).
         */
        connectBeginShape(shape: Shape, connectionSite: number): void;

        /**
         * Attaches the end of the specified connector to a specified shape.
         * @param shape The shape to connect.
         * @param connectionSite The connection site on the shape to which the end of the connector is attached. Must be an integer between 0 (inclusive) and the connection-site count of the specified shape (exclusive).
         */
        connectEndShape(shape: Shape, connectionSite: number): void;

        /**
         * Detaches the beginning of the specified connector from a shape.
         */
        disconnectBeginShape(): void;

        /**
         * Detaches the end of the specified connector from a shape.
         */
        disconnectEndShape(): void;
    }

    /**
     * Represents the fill formatting of a shape object.
     */
    interface ShapeFill {
        /**
         * Represents the shape fill foreground color in HTML color format, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange")
         */
        getForegroundColor(): string;

        /**
         * Represents the shape fill foreground color in HTML color format, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange")
         */
        setForegroundColor(foregroundColor: string): void;

        /**
         * Specifies the transparency percentage of the fill as a value from 0.0 (opaque) through 1.0 (clear). Returns `null` if the shape type does not support transparency or the shape fill has inconsistent transparency, such as with a gradient fill type.
         */
        getTransparency(): number;

        /**
         * Specifies the transparency percentage of the fill as a value from 0.0 (opaque) through 1.0 (clear). Returns `null` if the shape type does not support transparency or the shape fill has inconsistent transparency, such as with a gradient fill type.
         */
        setTransparency(transparency: number): void;

        /**
         * Returns the fill type of the shape. See `ExcelScript.ShapeFillType` for details.
         */
        getType(): ShapeFillType;

        /**
         * Clears the fill formatting of this shape.
         */
        clear(): void;

        /**
         * Sets the fill formatting of the shape to a uniform color. This changes the fill type to "Solid".
         * @param color A string that represents the fill color in HTML color format, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
         */
        setSolidColor(color: string): void;
    }

    /**
     * Represents the line formatting for the shape object. For images and geometric shapes, line formatting represents the border of the shape.
     */
    interface ShapeLineFormat {
        /**
         * Represents the line color in HTML color format, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
         */
        getColor(): string;

        /**
         * Represents the line color in HTML color format, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
         */
        setColor(color: string): void;

        /**
         * Represents the line style of the shape. Returns `null` when the line is not visible or there are inconsistent dash styles. See `ExcelScript.ShapeLineDashStyle` for details.
         */
        getDashStyle(): ShapeLineDashStyle;

        /**
         * Represents the line style of the shape. Returns `null` when the line is not visible or there are inconsistent dash styles. See `ExcelScript.ShapeLineDashStyle` for details.
         */
        setDashStyle(dashStyle: ShapeLineDashStyle): void;

        /**
         * Represents the line style of the shape. Returns `null` when the line is not visible or there are inconsistent styles. See `ExcelScript.ShapeLineStyle` for details.
         */
        getStyle(): ShapeLineStyle;

        /**
         * Represents the line style of the shape. Returns `null` when the line is not visible or there are inconsistent styles. See `ExcelScript.ShapeLineStyle` for details.
         */
        setStyle(style: ShapeLineStyle): void;

        /**
         * Represents the degree of transparency of the specified line as a value from 0.0 (opaque) through 1.0 (clear). Returns `null` when the shape has inconsistent transparencies.
         */
        getTransparency(): number;

        /**
         * Represents the degree of transparency of the specified line as a value from 0.0 (opaque) through 1.0 (clear). Returns `null` when the shape has inconsistent transparencies.
         */
        setTransparency(transparency: number): void;

        /**
         * Specifies if the line formatting of a shape element is visible. Returns `null` when the shape has inconsistent visibilities.
         */
        getVisible(): boolean;

        /**
         * Specifies if the line formatting of a shape element is visible. Returns `null` when the shape has inconsistent visibilities.
         */
        setVisible(visible: boolean): void;

        /**
         * Represents the weight of the line, in points. Returns `null` when the line is not visible or there are inconsistent line weights.
         */
        getWeight(): number;

        /**
         * Represents the weight of the line, in points. Returns `null` when the line is not visible or there are inconsistent line weights.
         */
        setWeight(weight: number): void;
    }

    /**
     * Represents the text frame of a shape object.
     */
    interface TextFrame {
        /**
         * The automatic sizing settings for the text frame. A text frame can be set to automatically fit the text to the text frame, to automatically fit the text frame to the text, or not perform any automatic sizing.
         */
        getAutoSizeSetting(): ShapeAutoSize;

        /**
         * The automatic sizing settings for the text frame. A text frame can be set to automatically fit the text to the text frame, to automatically fit the text frame to the text, or not perform any automatic sizing.
         */
        setAutoSizeSetting(autoSizeSetting: ShapeAutoSize): void;

        /**
         * Represents the bottom margin, in points, of the text frame.
         */
        getBottomMargin(): number;

        /**
         * Represents the bottom margin, in points, of the text frame.
         */
        setBottomMargin(bottomMargin: number): void;

        /**
         * Specifies if the text frame contains text.
         */
        getHasText(): boolean;

        /**
         * Represents the horizontal alignment of the text frame. See `ExcelScript.ShapeTextHorizontalAlignment` for details.
         */
        getHorizontalAlignment(): ShapeTextHorizontalAlignment;

        /**
         * Represents the horizontal alignment of the text frame. See `ExcelScript.ShapeTextHorizontalAlignment` for details.
         */
        setHorizontalAlignment(
            horizontalAlignment: ShapeTextHorizontalAlignment
        ): void;

        /**
         * Represents the horizontal overflow behavior of the text frame. See `ExcelScript.ShapeTextHorizontalOverflow` for details.
         */
        getHorizontalOverflow(): ShapeTextHorizontalOverflow;

        /**
         * Represents the horizontal overflow behavior of the text frame. See `ExcelScript.ShapeTextHorizontalOverflow` for details.
         */
        setHorizontalOverflow(
            horizontalOverflow: ShapeTextHorizontalOverflow
        ): void;

        /**
         * Represents the left margin, in points, of the text frame.
         */
        getLeftMargin(): number;

        /**
         * Represents the left margin, in points, of the text frame.
         */
        setLeftMargin(leftMargin: number): void;

        /**
         * Represents the angle to which the text is oriented for the text frame. See `ExcelScript.ShapeTextOrientation` for details.
         */
        getOrientation(): ShapeTextOrientation;

        /**
         * Represents the angle to which the text is oriented for the text frame. See `ExcelScript.ShapeTextOrientation` for details.
         */
        setOrientation(orientation: ShapeTextOrientation): void;

        /**
         * Represents the reading order of the text frame, either left-to-right or right-to-left. See `ExcelScript.ShapeTextReadingOrder` for details.
         */
        getReadingOrder(): ShapeTextReadingOrder;

        /**
         * Represents the reading order of the text frame, either left-to-right or right-to-left. See `ExcelScript.ShapeTextReadingOrder` for details.
         */
        setReadingOrder(readingOrder: ShapeTextReadingOrder): void;

        /**
         * Represents the right margin, in points, of the text frame.
         */
        getRightMargin(): number;

        /**
         * Represents the right margin, in points, of the text frame.
         */
        setRightMargin(rightMargin: number): void;

        /**
         * Represents the text that is attached to a shape in the text frame, and properties and methods for manipulating the text. See `ExcelScript.TextRange` for details.
         */
        getTextRange(): TextRange;

        /**
         * Represents the top margin, in points, of the text frame.
         */
        getTopMargin(): number;

        /**
         * Represents the top margin, in points, of the text frame.
         */
        setTopMargin(topMargin: number): void;

        /**
         * Represents the vertical alignment of the text frame. See `ExcelScript.ShapeTextVerticalAlignment` for details.
         */
        getVerticalAlignment(): ShapeTextVerticalAlignment;

        /**
         * Represents the vertical alignment of the text frame. See `ExcelScript.ShapeTextVerticalAlignment` for details.
         */
        setVerticalAlignment(
            verticalAlignment: ShapeTextVerticalAlignment
        ): void;

        /**
         * Represents the vertical overflow behavior of the text frame. See `ExcelScript.ShapeTextVerticalOverflow` for details.
         */
        getVerticalOverflow(): ShapeTextVerticalOverflow;

        /**
         * Represents the vertical overflow behavior of the text frame. See `ExcelScript.ShapeTextVerticalOverflow` for details.
         */
        setVerticalOverflow(verticalOverflow: ShapeTextVerticalOverflow): void;

        /**
         * Deletes all the text in the text frame.
         */
        deleteText(): void;
    }

    /**
     * Contains the text that is attached to a shape, in addition to properties and methods for manipulating the text.
     */
    interface TextRange {
        /**
         * Returns a `ShapeFont` object that represents the font attributes for the text range.
         */
        getFont(): ShapeFont;

        /**
         * Represents the plain text content of the text range.
         */
        getText(): string;

        /**
         * Represents the plain text content of the text range.
         */
        setText(text: string): void;

        /**
         * Returns a TextRange object for the substring in the given range.
         * @param start The zero-based index of the first character to get from the text range.
         * @param length Optional. The number of characters to be returned in the new text range. If length is omitted, all the characters from start to the end of the text range's last paragraph will be returned.
         */
        getSubstring(start: number, length?: number): TextRange;
    }

    /**
     * Represents the font attributes, such as font name, font size, and color, for a shape's `TextRange` object.
     */
    interface ShapeFont {
        /**
         * Represents the bold status of font. Returns `null` if the `TextRange` includes both bold and non-bold text fragments.
         */
        getBold(): boolean;

        /**
         * Represents the bold status of font. Returns `null` if the `TextRange` includes both bold and non-bold text fragments.
         */
        setBold(bold: boolean): void;

        /**
         * HTML color code representation of the text color (e.g., "#FF0000" represents red). Returns `null` if the `TextRange` includes text fragments with different colors.
         */
        getColor(): string;

        /**
         * HTML color code representation of the text color (e.g., "#FF0000" represents red). Returns `null` if the `TextRange` includes text fragments with different colors.
         */
        setColor(color: string): void;

        /**
         * Represents the italic status of font. Returns `null` if the `TextRange` includes both italic and non-italic text fragments.
         */
        getItalic(): boolean;

        /**
         * Represents the italic status of font. Returns `null` if the `TextRange` includes both italic and non-italic text fragments.
         */
        setItalic(italic: boolean): void;

        /**
         * Represents font name (e.g., "Calibri"). If the text is a Complex Script or East Asian language, this is the corresponding font name; otherwise it is the Latin font name.
         */
        getName(): string;

        /**
         * Represents font name (e.g., "Calibri"). If the text is a Complex Script or East Asian language, this is the corresponding font name; otherwise it is the Latin font name.
         */
        setName(name: string): void;

        /**
         * Represents font size in points (e.g., 11). Returns `null` if the `TextRange` includes text fragments with different font sizes.
         */
        getSize(): number;

        /**
         * Represents font size in points (e.g., 11). Returns `null` if the `TextRange` includes text fragments with different font sizes.
         */
        setSize(size: number): void;

        /**
         * Type of underline applied to the font. Returns `null` if the `TextRange` includes text fragments with different underline styles. See `ExcelScript.ShapeFontUnderlineStyle` for details.
         */
        getUnderline(): ShapeFontUnderlineStyle;

        /**
         * Type of underline applied to the font. Returns `null` if the `TextRange` includes text fragments with different underline styles. See `ExcelScript.ShapeFontUnderlineStyle` for details.
         */
        setUnderline(underline: ShapeFontUnderlineStyle): void;
    }

    /**
     * Represents a `Slicer` object in the workbook.
     */
    interface Slicer {
        /**
         * Represents the caption of the slicer.
         */
        getCaption(): string;

        /**
         * Represents the caption of the slicer.
         */
        setCaption(caption: string): void;

        /**
         * Represents the height, in points, of the slicer.
         * Throws an `InvalidArgument` exception when set with a negative value or zero as an input.
         */
        getHeight(): number;

        /**
         * Represents the height, in points, of the slicer.
         * Throws an `InvalidArgument` exception when set with a negative value or zero as an input.
         */
        setHeight(height: number): void;

        /**
         * Represents the unique ID of the slicer.
         */
        getId(): string;

        /**
         * Value is `true` if all filters currently applied on the slicer are cleared.
         */
        getIsFilterCleared(): boolean;

        /**
         * Represents the distance, in points, from the left side of the slicer to the left of the worksheet.
         * Throws an `InvalidArgument` error when set with a negative value as an input.
         */
        getLeft(): number;

        /**
         * Represents the distance, in points, from the left side of the slicer to the left of the worksheet.
         * Throws an `InvalidArgument` error when set with a negative value as an input.
         */
        setLeft(left: number): void;

        /**
         * Represents the name of the slicer.
         */
        getName(): string;

        /**
         * Represents the name of the slicer.
         */
        setName(name: string): void;

        /**
         * Represents the sort order of the items in the slicer. Possible values are: "DataSourceOrder", "Ascending", "Descending".
         */
        getSortBy(): SlicerSortType;

        /**
         * Represents the sort order of the items in the slicer. Possible values are: "DataSourceOrder", "Ascending", "Descending".
         */
        setSortBy(sortBy: SlicerSortType): void;

        /**
         * Constant value that represents the slicer style. Possible values are: "SlicerStyleLight1" through "SlicerStyleLight6", "TableStyleOther1" through "TableStyleOther2", "SlicerStyleDark1" through "SlicerStyleDark6". A custom user-defined style present in the workbook can also be specified.
         */
        getStyle(): string;

        /**
         * Constant value that represents the slicer style. Possible values are: "SlicerStyleLight1" through "SlicerStyleLight6", "TableStyleOther1" through "TableStyleOther2", "SlicerStyleDark1" through "SlicerStyleDark6". A custom user-defined style present in the workbook can also be specified.
         */
        setStyle(style: string): void;

        /**
         * Represents the distance, in points, from the top edge of the slicer to the top of the worksheet.
         * Throws an `InvalidArgument` error when set with a negative value as an input.
         */
        getTop(): number;

        /**
         * Represents the distance, in points, from the top edge of the slicer to the top of the worksheet.
         * Throws an `InvalidArgument` error when set with a negative value as an input.
         */
        setTop(top: number): void;

        /**
         * Represents the width, in points, of the slicer.
         * Throws an `InvalidArgument` error when set with a negative value or zero as an input.
         */
        getWidth(): number;

        /**
         * Represents the width, in points, of the slicer.
         * Throws an `InvalidArgument` error when set with a negative value or zero as an input.
         */
        setWidth(width: number): void;

        /**
         * Represents the worksheet containing the slicer.
         */
        getWorksheet(): Worksheet;

        /**
         * Clears all the filters currently applied on the slicer.
         */
        clearFilters(): void;

        /**
         * Deletes the slicer.
         */
        delete(): void;

        /**
         * Returns an array of selected items' keys.
         */
        getSelectedItems(): string[];

        /**
         * Selects slicer items based on their keys. The previous selections are cleared.
         * All items will be selected by default if the array is empty.
         * @param items Optional. The specified slicer item names to be selected.
         */
        selectItems(items?: string[]): void;

        /**
         * Represents the collection of slicer items that are part of the slicer.
         */
        getSlicerItems(): SlicerItem[];

        /**
         * Gets a slicer item using its key or name. If the slicer item doesn't exist, then this method returns `undefined`.
         * @param key Key or name of the slicer to be retrieved.
         */
        getSlicerItem(key: string): SlicerItem | undefined;
    }

    /**
     * Represents a slicer item in a slicer.
     */
    interface SlicerItem {
        /**
         * Value is `true` if the slicer item has data.
         */
        getHasData(): boolean;

        /**
         * Value is `true` if the slicer item is selected.
         * Setting this value will not clear the selected state of other slicer items.
         * By default, if the slicer item is the only one selected, when it is deselected, all items will be selected.
         */
        getIsSelected(): boolean;

        /**
         * Value is `true` if the slicer item is selected.
         * Setting this value will not clear the selected state of other slicer items.
         * By default, if the slicer item is the only one selected, when it is deselected, all items will be selected.
         */
        setIsSelected(isSelected: boolean): void;

        /**
         * Represents the unique value representing the slicer item.
         */
        getKey(): string;

        /**
         * Represents the title displayed in the Excel UI.
         */
        getName(): string;
    }

    /**
     * Represents a named sheet view of a worksheet. A sheet view stores the sort and filter rules for a particular worksheet.
     * Every sheet view (even a temporary sheet view) has a unique, worksheet-scoped name that is used to access the view.
     */
    interface NamedSheetView {
        /**
         * Gets or sets the name of the sheet view.
         * The temporary sheet view name is the empty string ("").  Naming the view by using the name property causes the sheet view to be saved.
         */
        getName(): string;

        /**
         * Gets or sets the name of the sheet view.
         * The temporary sheet view name is the empty string ("").  Naming the view by using the name property causes the sheet view to be saved.
         */
        setName(name: string): void;

        /**
         * Activates this sheet view. This is equivalent to using "Switch To" in the Excel UI.
         */
        activate(): void;

        /**
         * Removes the sheet view from the worksheet.
         */
        delete(): void;

        /**
         * Creates a copy of this sheet view.
         * @param name The name of the duplicated sheet view. If no name is provided, one will be generated.
         */
        duplicate(name?: string): NamedSheetView;
    }

    //
    // Interface
    //

    /**
     * Configurable template for a date filter to apply to a PivotField.
     * The `condition` defines what criteria need to be set in order for the filter to operate.
     */
    interface PivotDateFilter {
        /**
         * The comparator is the static value to which other values are compared. The type of comparison is defined by the condition.
         */
        comparator?: FilterDatetime;

        /**
         * Specifies the condition for the filter, which defines the necessary filtering criteria.
         */
        condition: DateFilterCondition;

        /**
         * If `true`, filter *excludes* items that meet criteria. The default is `false` (filter to include items that meet criteria).
         */
        exclusive?: boolean;

        /**
         * The lower-bound of the range for the `between` filter condition.
         */
        lowerBound?: FilterDatetime;

        /**
         * The upper-bound of the range for the `between` filter condition.
         */
        upperBound?: FilterDatetime;

        /**
         * For `equals`, `before`, `after`, and `between` filter conditions, indicates if comparisons should be made as whole days.
         */
        wholeDays?: boolean;
    }

    /**
     * An interface representing all PivotFilters currently applied to a given PivotField.
     */
    interface PivotFilters {
        /**
         * The PivotField's currently applied date filter. This property is `null` if no value filter is applied.
         */
        dateFilter?: PivotDateFilter;

        /**
         * The PivotField's currently applied label filter. This property is `null` if no value filter is applied.
         */
        labelFilter?: PivotLabelFilter;

        /**
         * The PivotField's currently applied manual filter. This property is `null` if no value filter is applied.
         */
        manualFilter?: PivotManualFilter;

        /**
         * The PivotField's currently applied value filter. This property is `null` if no value filter is applied.
         */
        valueFilter?: PivotValueFilter;
    }

    /**
     * Configurable template for a label filter to apply to a PivotField.
     * The `condition` defines what criteria need to be set in order for the filter to operate.
     */
    interface PivotLabelFilter {
        /**
         * Specifies the condition for the filter, which defines the necessary filtering criteria.
         */
        condition: LabelFilterCondition;

        /**
         * If `true`, filter *excludes* items that meet criteria. The default is `false` (filter to include items that meet criteria).
         */
        exclusive?: boolean;

        /**
         * The lower-bound of the range for the `between` filter condition.
         * Note: A numeric string is treated as a number when being compared against other numeric strings.
         */
        lowerBound?: string;

        /**
         * The substring used for `beginsWith`, `endsWith`, and `contains` filter conditions.
         */
        substring?: string;

        /**
         * The upper-bound of the range for the `between` filter condition.
         * Note: A numeric string is treated as a number when being compared against other numeric strings.
         */
        upperBound?: string;
    }

    /**
     * Configurable template for a manual filter to apply to a PivotField.
     * The `condition` defines what criteria need to be set in order for the filter to operate.
     */
    interface PivotManualFilter {
        /**
         * A list of selected items to manually filter. These must be existing and valid items from the chosen field.
         */
        selectedItems?: (string | PivotItem)[];
    }

    /**
     * Configurable template for a value filter to apply to a PivotField.
     * The `condition` defines what criteria need to be set in order for the filter to operate.
     */
    interface PivotValueFilter {
        /**
         * The comparator is the static value to which other values are compared. The type of comparison is defined by the condition.
         * For example, if comparator is "50" and condition is "greaterThan", all item values that are not greater than 50 will be removed by the filter.
         */
        comparator?: number;

        /**
         * Specifies the condition for the filter, which defines the necessary filtering criteria.
         */
        condition: ValueFilterCondition;

        /**
         * If `true`, filter *excludes* items that meet criteria. The default is `false` (filter to include items that meet criteria).
         */
        exclusive?: boolean;

        /**
         * The lower-bound of the range for the `between` filter condition.
         */
        lowerBound?: number;

        /**
         * Specifies if the filter is for the top/bottom N items, top/bottom N percent, or top/bottom N sum.
         */
        selectionType?: TopBottomSelectionType;

        /**
         * The "N" threshold number of items, percent, or sum to be filtered for a top/bottom filter condition.
         */
        threshold?: number;

        /**
         * The upper-bound of the range for the `between` filter condition.
         */
        upperBound?: number;

        /**
         * Name of the chosen "value" in the field by which to filter.
         */
        value: string;
    }

    /**
     * Represents the options in sheet protection.
     */
    interface WorksheetProtectionOptions {
        /**
         * Represents the worksheet protection option allowing use of the AutoFilter feature.
         */
        allowAutoFilter?: boolean;

        /**
         * Represents the worksheet protection option allowing deleting of columns.
         */
        allowDeleteColumns?: boolean;

        /**
         * Represents the worksheet protection option allowing deleting of rows.
         */
        allowDeleteRows?: boolean;

        /**
         * Represents the worksheet protection option allowing editing of objects.
         */
        allowEditObjects?: boolean;

        /**
         * Represents the worksheet protection option allowing editing of scenarios.
         */
        allowEditScenarios?: boolean;

        /**
         * Represents the worksheet protection option allowing formatting of cells.
         */
        allowFormatCells?: boolean;

        /**
         * Represents the worksheet protection option allowing formatting of columns.
         */
        allowFormatColumns?: boolean;

        /**
         * Represents the worksheet protection option allowing formatting of rows.
         */
        allowFormatRows?: boolean;

        /**
         * Represents the worksheet protection option allowing inserting of columns.
         */
        allowInsertColumns?: boolean;

        /**
         * Represents the worksheet protection option allowing inserting of hyperlinks.
         */
        allowInsertHyperlinks?: boolean;

        /**
         * Represents the worksheet protection option allowing inserting of rows.
         */
        allowInsertRows?: boolean;

        /**
         * Represents the worksheet protection option allowing use of the PivotTable feature.
         */
        allowPivotTables?: boolean;

        /**
         * Represents the worksheet protection option allowing use of the sort feature.
         */
        allowSort?: boolean;

        /**
         * Represents the worksheet protection option of selection mode.
         */
        selectionMode?: ProtectionSelectionMode;
    }

    /**
     * Represents the necessary strings to get/set a hyperlink (XHL) object.
     */
    interface RangeHyperlink {
        /**
         * Represents the URL target for the hyperlink.
         */
        address?: string;

        /**
         * Represents the document reference target for the hyperlink.
         */
        documentReference?: string;

        /**
         * Represents the string displayed when hovering over the hyperlink.
         */
        screenTip?: string;

        /**
         * Represents the string that is displayed in the top left most cell in the range.
         */
        textToDisplay?: string;
    }

    /**
     * Represents the search criteria to be used.
     */
    interface SearchCriteria {
        /**
         * Specifies if the match needs to be complete or partial.
         * A complete match matches the entire contents of the cell. A partial match matches a substring within the content of the cell (e.g., `cat` partially matches `caterpillar` and `scatter`).
         * Default is `false` (partial).
         */
        completeMatch?: boolean;

        /**
         * Specifies if the match is case-sensitive. Default is `false` (case-insensitive).
         */
        matchCase?: boolean;

        /**
         * Specifies the search direction. Default is forward. See `ExcelScript.SearchDirection`.
         */
        searchDirection?: SearchDirection;
    }

    /**
     * Represents the worksheet search criteria to be used.
     */
    interface WorksheetSearchCriteria {
        /**
         * Specifies if the match needs to be complete or partial.
         * A complete match matches the entire contents of the cell. A partial match matches a substring within the content of the cell (e.g., `cat` partially matches `caterpillar` and `scatter`).
         * Default is `false` (partial).
         */
        completeMatch?: boolean;

        /**
         * Specifies if the match is case-sensitive. Default is `false` (case-insensitive).
         */
        matchCase?: boolean;
    }

    /**
     * Represents the replace criteria to be used.
     */
    interface ReplaceCriteria {
        /**
         * Specifies if the match needs to be complete or partial.
         * A complete match matches the entire contents of the cell. A partial match matches a substring within the content of the cell (e.g., `cat` partially matches `caterpillar` and `scatter`).
         * Default is `false` (partial).
         */
        completeMatch?: boolean;

        /**
         * Specifies if the match is case-sensitive. Default is `false` (case-insensitive).
         */
        matchCase?: boolean;
    }

    /**
     * A data validation rule contains different types of data validation. You can only use one of them at a time according the `ExcelScript.DataValidationType`.
     */
    interface DataValidationRule {
        /**
         * Custom data validation criteria.
         */
        custom?: CustomDataValidation;

        /**
         * Date data validation criteria.
         */
        date?: DateTimeDataValidation;

        /**
         * Decimal data validation criteria.
         */
        decimal?: BasicDataValidation;

        /**
         * List data validation criteria.
         */
        list?: ListDataValidation;

        /**
         * Text length data validation criteria.
         */
        textLength?: BasicDataValidation;

        /**
         * Time data validation criteria.
         */
        time?: DateTimeDataValidation;

        /**
         * Whole number data validation criteria.
         */
        wholeNumber?: BasicDataValidation;
    }

    /**
     * Represents the basic type data validation criteria.
     */
    interface BasicDataValidation {
        /**
         * Specifies the right-hand operand when the operator property is set to a binary operator such as GreaterThan (the left-hand operand is the value the user tries to enter in the cell). With the ternary operators Between and NotBetween, specifies the lower bound operand.
         * For example, setting formula1 to 10 and operator to GreaterThan means that valid data for the range must be greater than 10.
         * When setting the value, it can be passed in as a number, a range object, or a string formula (where the string is either a stringified number, a cell reference like "=A1", or a formula like "=MIN(A1, B1)").
         * When retrieving the value, it will always be returned as a string formula, for example: "=10", "=A1", "=SUM(A1:B5)", etc.
         */
        formula1: string | number | Range;

        /**
         * With the ternary operators Between and NotBetween, specifies the upper bound operand. Is not used with the binary operators, such as GreaterThan.
         * When setting the value, it can be passed in as a number, a range object, or a string formula (where the string is either a stringified number, a cell reference like "=A1", or a formula like "=MIN(A1, B1)").
         * When retrieving the value, it will always be returned as a string formula, for example: "=10", "=A1", "=SUM(A1:B5)", etc.
         */
        formula2?: string | number | Range;

        /**
         * The operator to use for validating the data.
         */
        operator: DataValidationOperator;
    }

    /**
     * Represents the date data validation criteria.
     */
    interface DateTimeDataValidation {
        /**
         * Specifies the right-hand operand when the operator property is set to a binary operator such as GreaterThan (the left-hand operand is the value the user tries to enter in the cell). With the ternary operators Between and NotBetween, specifies the lower bound operand.
         * When setting the value, it can be passed in as a Date, a Range object, or a string formula (where the string is either a stringified date/time in ISO8601 format, a cell reference like "=A1", or a formula like "=MIN(A1, B1)").
         * When retrieving the value, it will always be returned as a string formula, for example: "=10", "=A1", "=SUM(A1:B5)", etc.
         */
        formula1: string | Date | Range;

        /**
         * With the ternary operators Between and NotBetween, specifies the upper bound operand. Is not used with the binary operators, such as GreaterThan.
         * When setting the value, it can be passed in as a Date, a Range object, or a string (where the string is either a stringified date/time in ISO8601 format, a cell reference like "=A1", or a formula like "=MIN(A1, B1)").
         * When retrieving the value, it will always be returned as a string formula, for example: "=10", "=A1", "=SUM(A1:B5)", etc.
         */
        formula2?: string | Date | Range;

        /**
         * The operator to use for validating the data.
         */
        operator: DataValidationOperator;
    }

    /**
     * Represents the List data validation criteria.
     */
    interface ListDataValidation {
        /**
         * Specifies whether to display the list in a cell drop-down. The default is `true`.
         */
        inCellDropDown: boolean;

        /**
         * Source of the list for data validation
         * When setting the value, it can be passed in as a `Range` object, or a string that contains a comma-separated number, boolean, or date.
         */
        source: string | Range;
    }

    /**
     * Represents the custom data validation criteria.
     */
    interface CustomDataValidation {
        /**
         * A custom data validation formula. This creates special input rules, such as preventing duplicates, or limiting the total in a range of cells.
         */
        formula: string;
    }

    /**
     * Represents the error alert properties for the data validation.
     */
    interface DataValidationErrorAlert {
        /**
         * Represents the error alert message.
         */
        message: string;

        /**
         * Specifies whether to show an error alert dialog when a user enters invalid data. The default is `true`.
         */
        showAlert: boolean;

        /**
         * The data validation alert type, please see `ExcelScript.DataValidationAlertStyle` for details.
         */
        style: DataValidationAlertStyle;

        /**
         * Represents the error alert dialog title.
         */
        title: string;
    }

    /**
     * Represents the user prompt properties for the data validation.
     */
    interface DataValidationPrompt {
        /**
         * Specifies the message of the prompt.
         */
        message: string;

        /**
         * Specifies if a prompt is shown when a user selects a cell with data validation.
         */
        showPrompt: boolean;

        /**
         * Specifies the title for the prompt.
         */
        title: string;
    }

    /**
     * Represents a condition in a sorting operation.
     */
    interface SortField {
        /**
         * Specifies if the sorting is done in an ascending fashion.
         */
        ascending?: boolean;

        /**
         * Specifies the color that is the target of the condition if the sorting is on font or cell color.
         */
        color?: string;

        /**
         * Represents additional sorting options for this field.
         */
        dataOption?: SortDataOption;

        /**
         * Specifies the icon that is the target of the condition, if the sorting is on the cell's icon.
         */
        icon?: Icon;

        /**
         * Specifies the column (or row, depending on the sort orientation) that the condition is on. Represented as an offset from the first column (or row).
         */
        key: number;

        /**
         * Specifies the type of sorting of this condition.
         */
        sortOn?: SortOn;

        /**
         * Specifies the subfield that is the target property name of a rich value to sort on.
         */
        subField?: string;
    }

    /**
     * Represents the filtering criteria applied to a column.
     */
    interface FilterCriteria {
        /**
         * The HTML color string used to filter cells. Used with `cellColor` and `fontColor` filtering.
         */
        color?: string;

        /**
         * The first criterion used to filter data. Used as an operator in the case of `custom` filtering.
         * For example ">50" for numbers greater than 50, or "=*s" for values ending in "s".
         *
         * Used as a number in the case of top/bottom items/percents (e.g., "5" for the top 5 items if `filterOn` is set to `topItems`).
         */
        criterion1?: string;

        /**
         * The second criterion used to filter data. Only used as an operator in the case of `custom` filtering.
         */
        criterion2?: string;

        /**
         * The dynamic criteria from the `ExcelScript.DynamicFilterCriteria` set to apply on this column. Used with `dynamic` filtering.
         */
        dynamicCriteria?: DynamicFilterCriteria;

        /**
         * The property used by the filter to determine whether the values should stay visible.
         */
        filterOn: FilterOn;

        /**
         * The icon used to filter cells. Used with `icon` filtering.
         */
        icon?: Icon;

        /**
         * The operator used to combine criterion 1 and 2 when using `custom` filtering.
         */
        operator?: FilterOperator;

        /**
         * The property used by the filter to do a rich filter on rich values.
         */
        subField?: string;

        /**
         * The set of values to be used as part of `values` filtering.
         */
        values?: Array<string | FilterDatetime>;
    }

    /**
     * Represents how to filter a date when filtering on values.
     */
    interface FilterDatetime {
        /**
         * The date in ISO8601 format used to filter data.
         */
        date: string;

        /**
         * How specific the date should be used to keep data. For example, if the date is 2005-04-02 and the specifity is set to "month", the filter operation will keep all rows with a date in the month of April 2005.
         */
        specificity: FilterDatetimeSpecificity;
    }

    /**
     * Represents a cell icon.
     */
    interface Icon {
        /**
         * Specifies the index of the icon in the given set.
         */
        index: number;

        /**
         * Specifies the set that the icon is part of.
         */
        set: IconSet;
    }

    interface ShowAsRule {
        /**
         * The PivotField to base the `ShowAs` calculation on, if applicable according to the `ShowAsCalculation` type, else `null`.
         */
        baseField?: PivotField;

        /**
         * The item to base the `ShowAs` calculation on, if applicable according to the `ShowAsCalculation` type, else `null`.
         */
        baseItem?: PivotItem;

        /**
         * The `ShowAs` calculation to use for the PivotField. See `ExcelScript.ShowAsCalculation` for details.
         */
        calculation: ShowAsCalculation;
    }

    /**
     * Subtotals for the Pivot Field.
     */
    interface Subtotals {
        /**
         * If `Automatic` is set to `true`, then all other values will be ignored when setting the `Subtotals`.
         */
        automatic?: boolean;

        /**
         * Average
         */
        average?: boolean;

        /**
         * Count
         */
        count?: boolean;

        /**
         * CountNumbers
         */
        countNumbers?: boolean;

        /**
         * Max
         */
        max?: boolean;

        /**
         * Min
         */
        min?: boolean;

        /**
         * Product
         */
        product?: boolean;

        /**
         * StandardDeviation
         */
        standardDeviation?: boolean;

        /**
         * StandardDeviationP
         */
        standardDeviationP?: boolean;

        /**
         * Sum
         */
        sum?: boolean;

        /**
         * Variance
         */
        variance?: boolean;

        /**
         * VarianceP
         */
        varianceP?: boolean;
    }

    /**
     * Represents a rule-type for a data bar.
     */
    interface ConditionalDataBarRule {
        /**
         * The formula, if required, on which to evaluate the data bar rule.
         */
        formula?: string;

        /**
         * The type of rule for the data bar.
         */
        type: ConditionalFormatRuleType;
    }

    /**
     * Represents an icon criterion which contains a type, value, an operator, and an optional custom icon, if not using an icon set.
     */
    interface ConditionalIconCriterion {
        /**
         * The custom icon for the current criterion, if different from the default icon set, else `null` will be returned.
         */
        customIcon?: Icon;

        /**
         * A number or a formula depending on the type.
         */
        formula: string;

        /**
         * `greaterThan` or `greaterThanOrEqual` for each of the rule types for the icon conditional format.
         */
        operator: ConditionalIconCriterionOperator;

        /**
         * What the icon conditional formula should be based on.
         */
        type: ConditionalFormatIconRuleType;
    }

    /**
     * Represents the criteria of the color scale.
     */
    interface ConditionalColorScaleCriteria {
        /**
         * The maximum point of the color scale criterion.
         */
        maximum: ConditionalColorScaleCriterion;

        /**
         * The midpoint of the color scale criterion, if the color scale is a 3-color scale.
         */
        midpoint?: ConditionalColorScaleCriterion;

        /**
         * The minimum point of the color scale criterion.
         */
        minimum: ConditionalColorScaleCriterion;
    }

    /**
     * Represents a color scale criterion which contains a type, value, and a color.
     */
    interface ConditionalColorScaleCriterion {
        /**
         * HTML color code representation of the color scale color (e.g., #FF0000 represents Red).
         */
        color?: string;

        /**
         * A number, a formula, or `null` (if `type` is `lowestValue`).
         */
        formula?: string;

        /**
         * What the criterion conditional formula should be based on.
         */
        type: ConditionalFormatColorCriterionType;
    }

    /**
     * Represents the rule of the top/bottom conditional format.
     */
    interface ConditionalTopBottomRule {
        /**
         * The rank between 1 and 1000 for numeric ranks or 1 and 100 for percent ranks.
         */
        rank: number;

        /**
         * Format values based on the top or bottom rank.
         */
        type: ConditionalTopBottomCriterionType;
    }

    /**
     * Represents the preset criteria conditional format rule.
     */
    interface ConditionalPresetCriteriaRule {
        /**
         * The criterion of the conditional format.
         */
        criterion: ConditionalFormatPresetCriterion;
    }

    /**
     * Represents a cell value conditional format rule.
     */
    interface ConditionalTextComparisonRule {
        /**
         * The operator of the text conditional format.
         */
        operator: ConditionalTextOperator;

        /**
         * The text value of the conditional format.
         */
        text: string;
    }

    /**
     * Represents a cell value conditional format rule.
     */
    interface ConditionalCellValueRule {
        /**
         * The formula, if required, on which to evaluate the conditional format rule.
         */
        formula1: string;

        /**
         * The formula, if required, on which to evaluate the conditional format rule.
         */
        formula2?: string;

        /**
         * The operator of the cell value conditional format.
         */
        operator: ConditionalCellValueOperator;
    }

    /**
     * Represents page zoom properties.
     */
    interface PageLayoutZoomOptions {
        /**
         * Number of pages to fit horizontally. This value can be `null` if percentage scale is used.
         */
        horizontalFitToPages?: number;

        /**
         * Print page scale value can be between 10 and 400. This value can be `null` if fit to page tall or wide is specified.
         */
        scale?: number;

        /**
         * Number of pages to fit vertically. This value can be `null` if percentage scale is used.
         */
        verticalFitToPages?: number;
    }

    /**
     * Represents the options in page layout margins.
     */
    interface PageLayoutMarginOptions {
        /**
         * Specifies the page layout bottom margin in the unit specified to use for printing.
         */
        bottom?: number;

        /**
         * Specifies the page layout footer margin in the unit specified to use for printing.
         */
        footer?: number;

        /**
         * Specifies the page layout header margin in the unit specified to use for printing.
         */
        header?: number;

        /**
         * Specifies the page layout left margin in the unit specified to use for printing.
         */
        left?: number;

        /**
         * Specifies the page layout right margin in the unit specified to use for printing.
         */
        right?: number;

        /**
         * Specifies the page layout top margin in the unit specified to use for printing.
         */
        top?: number;
    }

    /**
     * Represents the entity that is mentioned in comments.
     */
    interface CommentMention {
        /**
         * The email address of the entity that is mentioned in a comment.
         */
        email: string;

        /**
         * The ID of the entity. The ID matches one of the IDs in `CommentRichContent.richContent`.
         */
        id: number;

        /**
         * The name of the entity that is mentioned in a comment.
         */
        name: string;
    }

    /**
     * Represents the content contained within a comment or comment reply. Rich content incudes the text string and any other objects contained within the comment body, such as mentions.
     */
    interface CommentRichContent {
        /**
         * An array containing all the entities (e.g., people) mentioned within the comment.
         */
        mentions?: CommentMention[];

        /**
         * Specifies the rich content of the comment (e.g., comment content with mentions, the first mentioned entity has an ID attribute of 0, and the second mentioned entity has an ID attribute of 1).
         */
        richContent: string;
    }

    //
    // Enum
    //

    /**
     * Represents the refresh mode of the workbook links.
     */
    enum WorkbookLinksRefreshMode {
        /**
         * The workbook links are updated manually.
         */
        manual,

        /**
         * The workbook links are updated at a set interval determined by the Excel application.
         */
        automatic,
    }

    /**
     * Enum representing all accepted conditions by which a date filter can be applied.
     * Used to configure the type of PivotFilter that is applied to the field.
     */
    enum DateFilterCondition {
        /**
         * `DateFilterCondition` is unknown or unsupported.
         */
        unknown,

        /**
         * Equals comparator criterion.
         *
         * Required Criteria: {`comparator`}.
         * Optional Criteria: {`wholeDays`, `exclusive`}.
         */
        equals,

        /**
         * Date is before comparator date.
         *
         * Required Criteria: {`comparator`}.
         * Optional Criteria: {`wholeDays`}.
         */
        before,

        /**
         * Date is before or equal to comparator date.
         *
         * Required Criteria: {`comparator`}.
         * Optional Criteria: {`wholeDays`}.
         */
        beforeOrEqualTo,

        /**
         * Date is after comparator date.
         *
         * Required Criteria: {`comparator`}.
         * Optional Criteria: {`wholeDays`}.
         */
        after,

        /**
         * Date is after or equal to comparator date.
         *
         * Required Criteria: {`comparator`}.
         * Optional Criteria: {`wholeDays`}.
         */
        afterOrEqualTo,

        /**
         * Between `lowerBound` and `upperBound` dates.
         *
         * Required Criteria: {`lowerBound`, `upperBound`}.
         * Optional Criteria: {`wholeDays`, `exclusive`}.
         */
        between,

        /**
         * Date is tomorrow.
         */
        tomorrow,

        /**
         * Date is today.
         */
        today,

        /**
         * Date is yesterday.
         */
        yesterday,

        /**
         * Date is next week.
         */
        nextWeek,

        /**
         * Date is this week.
         */
        thisWeek,

        /**
         * Date is last week.
         */
        lastWeek,

        /**
         * Date is next month.
         */
        nextMonth,

        /**
         * Date is this month.
         */
        thisMonth,

        /**
         * Date is last month.
         */
        lastMonth,

        /**
         * Date is next quarter.
         */
        nextQuarter,

        /**
         * Date is this quarter.
         */
        thisQuarter,

        /**
         * Date is last quarter.
         */
        lastQuarter,

        /**
         * Date is next year.
         */
        nextYear,

        /**
         * Date is this year.
         */
        thisYear,

        /**
         * Date is last year.
         */
        lastYear,

        /**
         * Date is in the same year to date.
         */
        yearToDate,

        /**
         * Date is in Quarter 1.
         */
        allDatesInPeriodQuarter1,

        /**
         * Date is in Quarter 2.
         */
        allDatesInPeriodQuarter2,

        /**
         * Date is in Quarter 3.
         */
        allDatesInPeriodQuarter3,

        /**
         * Date is in Quarter 4.
         */
        allDatesInPeriodQuarter4,

        /**
         * Date is in January.
         */
        allDatesInPeriodJanuary,

        /**
         * Date is in February.
         */
        allDatesInPeriodFebruary,

        /**
         * Date is in March.
         */
        allDatesInPeriodMarch,

        /**
         * Date is in April.
         */
        allDatesInPeriodApril,

        /**
         * Date is in May.
         */
        allDatesInPeriodMay,

        /**
         * Date is in June.
         */
        allDatesInPeriodJune,

        /**
         * Date is in July.
         */
        allDatesInPeriodJuly,

        /**
         * Date is in August.
         */
        allDatesInPeriodAugust,

        /**
         * Date is in September.
         */
        allDatesInPeriodSeptember,

        /**
         * Date is in October.
         */
        allDatesInPeriodOctober,

        /**
         * Date is in November.
         */
        allDatesInPeriodNovember,

        /**
         * Date is in December.
         */
        allDatesInPeriodDecember,
    }

    /**
     * Enum representing all accepted conditions by which a label filter can be applied.
     * Used to configure the type of PivotFilter that is applied to the field.
     * `PivotFilter.criteria.exclusive` can be set to `true` to invert many of these conditions.
     */
    enum LabelFilterCondition {
        /**
         * `LabelFilterCondition` is unknown or unsupported.
         */
        unknown,

        /**
         * Equals comparator criterion.
         *
         * Required Criteria: {`comparator`}.
         * Optional Criteria: {`exclusive`}.
         */
        equals,

        /**
         * Label begins with substring criterion.
         *
         * Required Criteria: {`substring`}.
         * Optional Criteria: {`exclusive`}.
         */
        beginsWith,

        /**
         * Label ends with substring criterion.
         *
         * Required Criteria: {`substring`}.
         * Optional Criteria: {`exclusive`}.
         */
        endsWith,

        /**
         * Label contains substring criterion.
         *
         * Required Criteria: {`substring`}.
         * Optional Criteria: {`exclusive`}.
         */
        contains,

        /**
         * Greater than comparator criterion.
         *
         * Required Criteria: {`comparator`}.
         */
        greaterThan,

        /**
         * Greater than or equal to comparator criterion.
         *
         * Required Criteria: {`comparator`}.
         */
        greaterThanOrEqualTo,

        /**
         * Less than comparator criterion.
         *
         * Required Criteria: {`comparator`}.
         */
        lessThan,

        /**
         * Less than or equal to comparator criterion.
         *
         * Required Criteria: {`comparator`}.
         */
        lessThanOrEqualTo,

        /**
         * Between `lowerBound` and `upperBound` criteria.
         *
         * Required Criteria: {`lowerBound`, `upperBound`}.
         * Optional Criteria: {`exclusive`}.
         */
        between,
    }

    /**
     * A simple enum that represents a type of filter for a PivotField.
     */
    enum PivotFilterType {
        /**
         * `PivotFilterType` is unknown or unsupported.
         */
        unknown,

        /**
         * Filters based on the value of a PivotItem with respect to a `DataPivotHierarchy`.
         */
        value,

        /**
         * Filters specific manually selected PivotItems from the PivotTable.
         */
        manual,

        /**
         * Filters PivotItems based on their labels.
         * Note: A PivotField cannot simultaneously have a label filter and a date filter applied.
         */
        label,

        /**
         * Filters PivotItems with a date in place of a label.
         * Note: A PivotField cannot simultaneously have a label filter and a date filter applied.
         */
        date,
    }

    /**
     * A simple enum for top/bottom filters to select whether to filter by the top N or bottom N percent, number, or sum of values.
     */
    enum TopBottomSelectionType {
        /**
         * Filter the top/bottom N number of items as measured by the chosen value.
         */
        items,

        /**
         * Filter the top/bottom N percent of items as measured by the chosen value.
         */
        percent,

        /**
         * Filter the top/bottom N sum as measured by the chosen value.
         */
        sum,
    }

    /**
     * Enum representing all accepted conditions by which a value filter can be applied.
     * Used to configure the type of PivotFilter that is applied to the field.
     * `PivotFilter.exclusive` can be set to `true` to invert many of these conditions.
     */
    enum ValueFilterCondition {
        /**
         * `ValueFilterCondition` is unknown or unsupported.
         */
        unknown,

        /**
         * Equals comparator criterion.
         *
         * Required Criteria: {`value`, `comparator`}.
         * Optional Criteria: {`exclusive`}.
         */
        equals,

        /**
         * Greater than comparator criterion.
         *
         * Required Criteria: {`value`, `comparator`}.
         */
        greaterThan,

        /**
         * Greater than or equal to comparator criterion.
         *
         * Required Criteria: {`value`, `comparator`}.
         */
        greaterThanOrEqualTo,

        /**
         * Less than comparator criterion.
         *
         * Required Criteria: {`value`, `comparator`}.
         */
        lessThan,

        /**
         * Less than or equal to comparator criterion.
         *
         * Required Criteria: {`value`, `comparator`}.
         */
        lessThanOrEqualTo,

        /**
         * Between `lowerBound` and `upperBound` criteria.
         *
         * Required Criteria: {`value`, `lowerBound`, `upperBound`}.
         * Optional Criteria: {`exclusive`}.
         */
        between,

        /**
         * In top N (`threshold`) [items, percent, sum] of value category.
         *
         * Required Criteria: {`value`, `threshold`, `selectionType`}.
         */
        topN,

        /**
         * In bottom N (`threshold`) [items, percent, sum] of value category.
         *
         * Required Criteria: {`value`, `threshold`, `selectionType`}.
         */
        bottomN,
    }

    /**
     * Represents the dimensions when getting values from chart series.
     */
    enum ChartSeriesDimension {
        /**
         * The chart series axis for the categories.
         */
        categories,

        /**
         * The chart series axis for the values.
         */
        values,

        /**
         * The chart series axis for the x-axis values in scatter and bubble charts.
         */
        xvalues,

        /**
         * The chart series axis for the y-axis values in scatter and bubble charts.
         */
        yvalues,

        /**
         * The chart series axis for the bubble sizes in bubble charts.
         */
        bubbleSizes,
    }

    /**
     * Represents the criteria for the top/bottom values filter.
     */
    enum PivotFilterTopBottomCriterion {
        invalid,

        topItems,

        topPercent,

        topSum,

        bottomItems,

        bottomPercent,

        bottomSum,
    }

    /**
     * Represents the sort direction.
     */
    enum SortBy {
        /**
         * Ascending sort. Smallest to largest or A to Z.
         */
        ascending,

        /**
         * Descending sort. Largest to smallest or Z to A.
         */
        descending,
    }

    /**
     * Aggregation function for the DataPivotField.
     */
    enum AggregationFunction {
        /**
         * Aggregation function is unknown or unsupported.
         */
        unknown,

        /**
         * Excel will automatically select the aggregation based on the data items.
         */
        automatic,

        /**
         * Aggregate using the sum of the data, equivalent to the SUM function.
         */
        sum,

        /**
         * Aggregate using the count of items in the data, equivalent to the COUNTA function.
         */
        count,

        /**
         * Aggregate using the average of the data, equivalent to the AVERAGE function.
         */
        average,

        /**
         * Aggregate using the maximum value of the data, equivalent to the MAX function.
         */
        max,

        /**
         * Aggregate using the minimum value of the data, equivalent to the MIN function.
         */
        min,

        /**
         * Aggregate using the product of the data, equivalent to the PRODUCT function.
         */
        product,

        /**
         * Aggregate using the count of numbers in the data, equivalent to the COUNT function.
         */
        countNumbers,

        /**
         * Aggregate using the standard deviation of the data, equivalent to the STDEV function.
         */
        standardDeviation,

        /**
         * Aggregate using the standard deviation of the data, equivalent to the STDEVP function.
         */
        standardDeviationP,

        /**
         * Aggregate using the variance of the data, equivalent to the VAR function.
         */
        variance,

        /**
         * Aggregate using the variance of the data, equivalent to the VARP function.
         */
        varianceP,
    }

    /**
     * The ShowAs calculation enum for the DataPivotField.
     */
    enum ShowAsCalculation {
        /**
         * Calculation is unknown or unsupported.
         */
        unknown,

        /**
         * No calculation is applied.
         */
        none,

        /**
         * Percent of the grand total.
         */
        percentOfGrandTotal,

        /**
         * Percent of the row total.
         */
        percentOfRowTotal,

        /**
         * Percent of the column total.
         */
        percentOfColumnTotal,

        /**
         * Percent of the row total for the specified Base field.
         */
        percentOfParentRowTotal,

        /**
         * Percent of the column total for the specified Base field.
         */
        percentOfParentColumnTotal,

        /**
         * Percent of the grand total for the specified Base field.
         */
        percentOfParentTotal,

        /**
         * Percent of the specified Base field and Base item.
         */
        percentOf,

        /**
         * Running total of the specified Base field.
         */
        runningTotal,

        /**
         * Percent running total of the specified Base field.
         */
        percentRunningTotal,

        /**
         * Difference from the specified Base field and Base item.
         */
        differenceFrom,

        /**
         * Difference from the specified Base field and Base item.
         */
        percentDifferenceFrom,

        /**
         * Ascending rank of the specified Base field.
         */
        rankAscending,

        /**
         * Descending rank of the specified Base field.
         */
        rankDecending,

        /**
         * Calculates the values as follows:
         * ((value in cell) x (Grand Total of Grand Totals)) / ((Grand Row Total) x (Grand Column Total))
         */
        index,
    }

    /**
     * Represents the axis from which to get the PivotItems.
     */
    enum PivotAxis {
        /**
         * The axis or region is unknown or unsupported.
         */
        unknown,

        /**
         * The row axis.
         */
        row,

        /**
         * The column axis.
         */
        column,

        /**
         * The data axis.
         */
        data,

        /**
         * The filter axis.
         */
        filter,
    }

    enum ChartAxisType {
        invalid,

        /**
         * Axis displays categories.
         */
        category,

        /**
         * Axis displays values.
         */
        value,

        /**
         * Axis displays data series.
         */
        series,
    }

    enum ChartAxisGroup {
        primary,

        secondary,
    }

    enum ChartAxisScaleType {
        linear,

        logarithmic,
    }

    enum ChartAxisPosition {
        automatic,

        maximum,

        minimum,

        custom,
    }

    enum ChartAxisTickMark {
        none,

        cross,

        inside,

        outside,
    }

    /**
     * Represents the state of calculation across the entire Excel application.
     */
    enum CalculationState {
        /**
         * Calculations complete.
         */
        done,

        /**
         * Calculations in progress.
         */
        calculating,

        /**
         * Changes that trigger calculation have been made, but a recalculation has not yet been performed.
         */
        pending,
    }

    enum ChartAxisTickLabelPosition {
        nextToAxis,

        high,

        low,

        none,
    }

    enum ChartAxisDisplayUnit {
        /**
         * Default option. This will reset display unit to the axis, and set unit label invisible.
         */
        none,

        /**
         * This will set the axis in units of hundreds.
         */
        hundreds,

        /**
         * This will set the axis in units of thousands.
         */
        thousands,

        /**
         * This will set the axis in units of tens of thousands.
         */
        tenThousands,

        /**
         * This will set the axis in units of hundreds of thousands.
         */
        hundredThousands,

        /**
         * This will set the axis in units of millions.
         */
        millions,

        /**
         * This will set the axis in units of tens of millions.
         */
        tenMillions,

        /**
         * This will set the axis in units of hundreds of millions.
         */
        hundredMillions,

        /**
         * This will set the axis in units of billions.
         */
        billions,

        /**
         * This will set the axis in units of trillions.
         */
        trillions,

        /**
         * This will set the axis in units of custom value.
         */
        custom,
    }

    /**
     * Specifies the unit of time for chart axes and data series.
     */
    enum ChartAxisTimeUnit {
        days,

        months,

        years,
    }

    /**
     * Represents the quartile calculation type of chart series layout. Only applies to a box and whisker chart.
     */
    enum ChartBoxQuartileCalculation {
        inclusive,

        exclusive,
    }

    /**
     * Specifies the type of the category axis.
     */
    enum ChartAxisCategoryType {
        /**
         * Excel controls the axis type.
         */
        automatic,

        /**
         * Axis groups data by an arbitrary set of categories.
         */
        textAxis,

        /**
         * Axis groups data on a time scale.
         */
        dateAxis,
    }

    /**
     * Specifies the bin type of a histogram chart or pareto chart series.
     */
    enum ChartBinType {
        category,

        auto,

        binWidth,

        binCount,
    }

    enum ChartLineStyle {
        none,

        continuous,

        dash,

        dashDot,

        dashDotDot,

        dot,

        grey25,

        grey50,

        grey75,

        automatic,

        roundDot,
    }

    enum ChartDataLabelPosition {
        invalid,

        none,

        center,

        insideEnd,

        insideBase,

        outsideEnd,

        left,

        right,

        top,

        bottom,

        bestFit,

        callout,
    }

    /**
     * Represents which parts of the error bar to include.
     */
    enum ChartErrorBarsInclude {
        both,

        minusValues,

        plusValues,
    }

    /**
     * Represents the range type for error bars.
     */
    enum ChartErrorBarsType {
        fixedValue,

        percent,

        stDev,

        stError,

        custom,
    }

    /**
     * Represents the mapping level of a chart series. This only applies to region map charts.
     */
    enum ChartMapAreaLevel {
        automatic,

        dataOnly,

        city,

        county,

        state,

        country,

        continent,

        world,
    }

    /**
     * Represents the gradient style of a chart series. This is only applicable for region map charts.
     */
    enum ChartGradientStyle {
        twoPhaseColor,

        threePhaseColor,
    }

    /**
     * Represents the gradient style type of a chart series. This is only applicable for region map charts.
     */
    enum ChartGradientStyleType {
        extremeValue,

        number,

        percent,
    }

    /**
     * Represents the position of the chart title.
     */
    enum ChartTitlePosition {
        automatic,

        top,

        bottom,

        left,

        right,
    }

    enum ChartLegendPosition {
        invalid,

        top,

        bottom,

        left,

        right,

        corner,

        custom,
    }

    enum ChartMarkerStyle {
        invalid,

        automatic,

        none,

        square,

        diamond,

        triangle,

        x,

        star,

        dot,

        dash,

        circle,

        plus,

        picture,
    }

    enum ChartPlotAreaPosition {
        automatic,

        custom,
    }

    /**
     * Represents the region level of a chart series layout. This only applies to region map charts.
     */
    enum ChartMapLabelStrategy {
        none,

        bestFit,

        showAll,
    }

    /**
     * Represents the region projection type of a chart series layout. This only applies to region map charts.
     */
    enum ChartMapProjectionType {
        automatic,

        mercator,

        miller,

        robinson,

        albers,
    }

    /**
     * Represents the parent label strategy of the chart series layout. This only applies to treemap charts
     */
    enum ChartParentLabelStrategy {
        none,

        banner,

        overlapping,
    }

    /**
     * Specifies whether the series are by rows or by columns. In Excel on desktop, the "auto" option will inspect the source data shape to automatically guess whether the data is by rows or columns. In Excel on the web, "auto" will simply default to "columns".
     */
    enum ChartSeriesBy {
        /**
         * In Excel on desktop, the "auto" option will inspect the source data shape to automatically guess whether the data is by rows or columns. In Excel on the web, "auto" will simply default to "columns".
         */
        auto,

        columns,

        rows,
    }

    /**
     * Represents the horizontal alignment for the specified object.
     */
    enum ChartTextHorizontalAlignment {
        center,

        left,

        right,

        justify,

        distributed,
    }

    /**
     * Represents the vertical alignment for the specified object.
     */
    enum ChartTextVerticalAlignment {
        center,

        bottom,

        top,

        justify,

        distributed,
    }

    enum ChartTickLabelAlignment {
        center,

        left,

        right,
    }

    enum ChartType {
        invalid,

        columnClustered,

        columnStacked,

        columnStacked100,

        barClustered,

        barStacked,

        barStacked100,

        lineStacked,

        lineStacked100,

        lineMarkers,

        lineMarkersStacked,

        lineMarkersStacked100,

        pieOfPie,

        pieExploded,

        barOfPie,

        xyscatterSmooth,

        xyscatterSmoothNoMarkers,

        xyscatterLines,

        xyscatterLinesNoMarkers,

        areaStacked,

        areaStacked100,

        doughnutExploded,

        radarMarkers,

        radarFilled,

        surface,

        surfaceWireframe,

        surfaceTopView,

        surfaceTopViewWireframe,

        bubble,

        bubble3DEffect,

        stockHLC,

        stockOHLC,

        stockVHLC,

        stockVOHLC,

        cylinderColClustered,

        cylinderColStacked,

        cylinderColStacked100,

        cylinderBarClustered,

        cylinderBarStacked,

        cylinderBarStacked100,

        cylinderCol,

        coneColClustered,

        coneColStacked,

        coneColStacked100,

        coneBarClustered,

        coneBarStacked,

        coneBarStacked100,

        coneCol,

        pyramidColClustered,

        pyramidColStacked,

        pyramidColStacked100,

        pyramidBarClustered,

        pyramidBarStacked,

        pyramidBarStacked100,

        pyramidCol,

        line,

        pie,

        xyscatter,

        area,

        doughnut,

        radar,

        histogram,

        boxwhisker,

        pareto,

        regionMap,

        treemap,

        waterfall,

        sunburst,

        funnel,
    }

    enum ChartUnderlineStyle {
        none,

        single,
    }

    enum ChartDisplayBlanksAs {
        notPlotted,

        zero,

        interplotted,
    }

    enum ChartPlotBy {
        rows,

        columns,
    }

    enum ChartSplitType {
        splitByPosition,

        splitByValue,

        splitByPercentValue,

        splitByCustomSplit,
    }

    enum ChartColorScheme {
        colorfulPalette1,

        colorfulPalette2,

        colorfulPalette3,

        colorfulPalette4,

        monochromaticPalette1,

        monochromaticPalette2,

        monochromaticPalette3,

        monochromaticPalette4,

        monochromaticPalette5,

        monochromaticPalette6,

        monochromaticPalette7,

        monochromaticPalette8,

        monochromaticPalette9,

        monochromaticPalette10,

        monochromaticPalette11,

        monochromaticPalette12,

        monochromaticPalette13,
    }

    enum ChartTrendlineType {
        linear,

        exponential,

        logarithmic,

        movingAverage,

        polynomial,

        power,
    }

    /**
     * Specifies where in the z-order a shape should be moved relative to other shapes.
     */
    enum ShapeZOrder {
        bringToFront,

        bringForward,

        sendToBack,

        sendBackward,
    }

    /**
     * Specifies the type of a shape.
     */
    enum ShapeType {
        unsupported,

        image,

        geometricShape,

        group,

        line,
    }

    /**
     * Specifies whether the shape is scaled relative to its original or current size.
     */
    enum ShapeScaleType {
        currentSize,

        originalSize,
    }

    /**
     * Specifies which part of the shape retains its position when the shape is scaled.
     */
    enum ShapeScaleFrom {
        scaleFromTopLeft,

        scaleFromMiddle,

        scaleFromBottomRight,
    }

    /**
     * Specifies a shape's fill type.
     */
    enum ShapeFillType {
        /**
         * No fill.
         */
        noFill,

        /**
         * Solid fill.
         */
        solid,

        /**
         * Gradient fill.
         */
        gradient,

        /**
         * Pattern fill.
         */
        pattern,

        /**
         * Picture and texture fill.
         */
        pictureAndTexture,

        /**
         * Mixed fill.
         */
        mixed,
    }

    /**
     * The type of underline applied to a font.
     */
    enum ShapeFontUnderlineStyle {
        none,

        single,

        double,

        heavy,

        dotted,

        dottedHeavy,

        dash,

        dashHeavy,

        dashLong,

        dashLongHeavy,

        dotDash,

        dotDashHeavy,

        dotDotDash,

        dotDotDashHeavy,

        wavy,

        wavyHeavy,

        wavyDouble,
    }

    /**
     * The format of the image.
     */
    enum PictureFormat {
        unknown,

        /**
         * Bitmap image.
         */
        bmp,

        /**
         * Joint Photographic Experts Group.
         */
        jpeg,

        /**
         * Graphics Interchange Format.
         */
        gif,

        /**
         * Portable Network Graphics.
         */
        png,

        /**
         * Scalable Vector Graphic.
         */
        svg,
    }

    /**
     * The style for a line.
     */
    enum ShapeLineStyle {
        /**
         * Single line.
         */
        single,

        /**
         * Thick line with a thin line on each side.
         */
        thickBetweenThin,

        /**
         * Thick line next to thin line. For horizontal lines, the thick line is above the thin line. For vertical lines, the thick line is to the left of the thin line.
         */
        thickThin,

        /**
         * Thick line next to thin line. For horizontal lines, the thick line is below the thin line. For vertical lines, the thick line is to the right of the thin line.
         */
        thinThick,

        /**
         * Two thin lines.
         */
        thinThin,
    }

    /**
     * The dash style for a line.
     */
    enum ShapeLineDashStyle {
        dash,

        dashDot,

        dashDotDot,

        longDash,

        longDashDot,

        roundDot,

        solid,

        squareDot,

        longDashDotDot,

        systemDash,

        systemDot,

        systemDashDot,
    }

    enum ArrowheadLength {
        short,

        medium,

        long,
    }

    enum ArrowheadStyle {
        none,

        triangle,

        stealth,

        diamond,

        oval,

        open,
    }

    enum ArrowheadWidth {
        narrow,

        medium,

        wide,
    }

    enum BindingType {
        range,

        table,

        text,
    }

    enum BorderIndex {
        edgeTop,

        edgeBottom,

        edgeLeft,

        edgeRight,

        insideVertical,

        insideHorizontal,

        diagonalDown,

        diagonalUp,
    }

    enum BorderLineStyle {
        none,

        continuous,

        dash,

        dashDot,

        dashDotDot,

        dot,

        double,

        slantDashDot,
    }

    enum BorderWeight {
        hairline,

        thin,

        medium,

        thick,
    }

    enum CalculationMode {
        /**
         * The default recalculation behavior where Excel calculates new formula results every time the relevant data is changed.
         */
        automatic,

        /**
         * Calculates new formula results every time the relevant data is changed, unless the formula is in a data table.
         */
        automaticExceptTables,

        /**
         * Calculations only occur when the user or add-in requests them.
         */
        manual,
    }

    enum CalculationType {
        /**
         * Recalculates all cells that Excel has marked as dirty, that is, dependents of volatile or changed data, and cells programmatically marked as dirty.
         */
        recalculate,

        /**
         * This will mark all cells as dirty and then recalculate them.
         */
        full,

        /**
         * This will rebuild the full dependency chain, mark all cells as dirty and then recalculate them.
         */
        fullRebuild,
    }

    enum ClearApplyTo {
        all,

        /**
         * Clears all formatting for the range.
         */
        formats,

        /**
         * Clears the contents of the range.
         */
        contents,

        /**
         * Clears all hyperlinks, but leaves all content and formatting intact.
         */
        hyperlinks,

        /**
         * Removes hyperlinks and formatting for the cell but leaves content, conditional formats, and data validation intact.
         */
        removeHyperlinks,
    }

    /**
     * Represents the format options for a data bar axis.
     */
    enum ConditionalDataBarAxisFormat {
        automatic,

        none,

        cellMidPoint,
    }

    /**
     * Represents the data bar direction within a cell.
     */
    enum ConditionalDataBarDirection {
        context,

        leftToRight,

        rightToLeft,
    }

    /**
     * Represents the direction for a selection.
     */
    enum ConditionalFormatDirection {
        top,

        bottom,
    }

    enum ConditionalFormatType {
        custom,

        dataBar,

        colorScale,

        iconSet,

        topBottom,

        presetCriteria,

        containsText,

        cellValue,
    }

    /**
     * Represents the types of conditional format values.
     */
    enum ConditionalFormatRuleType {
        invalid,

        automatic,

        lowestValue,

        highestValue,

        number,

        percent,

        formula,

        percentile,
    }

    /**
     * Represents the types of icon conditional format.
     */
    enum ConditionalFormatIconRuleType {
        invalid,

        number,

        percent,

        formula,

        percentile,
    }

    /**
     * Represents the types of color criterion for conditional formatting.
     */
    enum ConditionalFormatColorCriterionType {
        invalid,

        lowestValue,

        highestValue,

        number,

        percent,

        formula,

        percentile,
    }

    /**
     * Represents the criteria for the above/below average conditional format type.
     */
    enum ConditionalTopBottomCriterionType {
        invalid,

        topItems,

        topPercent,

        bottomItems,

        bottomPercent,
    }

    /**
     * Represents the criteria of the preset criteria conditional format type.
     */
    enum ConditionalFormatPresetCriterion {
        invalid,

        blanks,

        nonBlanks,

        errors,

        nonErrors,

        yesterday,

        today,

        tomorrow,

        lastSevenDays,

        lastWeek,

        thisWeek,

        nextWeek,

        lastMonth,

        thisMonth,

        nextMonth,

        aboveAverage,

        belowAverage,

        equalOrAboveAverage,

        equalOrBelowAverage,

        oneStdDevAboveAverage,

        oneStdDevBelowAverage,

        twoStdDevAboveAverage,

        twoStdDevBelowAverage,

        threeStdDevAboveAverage,

        threeStdDevBelowAverage,

        uniqueValues,

        duplicateValues,
    }

    /**
     * Represents the operator of the text conditional format type.
     */
    enum ConditionalTextOperator {
        invalid,

        contains,

        notContains,

        beginsWith,

        endsWith,
    }

    /**
     * Represents the operator of the text conditional format type.
     */
    enum ConditionalCellValueOperator {
        invalid,

        between,

        notBetween,

        equalTo,

        notEqualTo,

        greaterThan,

        lessThan,

        greaterThanOrEqual,

        lessThanOrEqual,
    }

    /**
     * Represents the operator for each icon criteria.
     */
    enum ConditionalIconCriterionOperator {
        invalid,

        greaterThan,

        greaterThanOrEqual,
    }

    enum ConditionalRangeBorderIndex {
        edgeTop,

        edgeBottom,

        edgeLeft,

        edgeRight,
    }

    enum ConditionalRangeBorderLineStyle {
        none,

        continuous,

        dash,

        dashDot,

        dashDotDot,

        dot,
    }

    enum ConditionalRangeFontUnderlineStyle {
        none,

        single,

        double,
    }

    /**
     * Represents the data validation type enum.
     */
    enum DataValidationType {
        /**
         * None means allow any value, indicating that there is no data validation in the range.
         */
        none,

        /**
         * The whole number data validation type.
         */
        wholeNumber,

        /**
         * The decimal data validation type.
         */
        decimal,

        /**
         * The list data validation type.
         */
        list,

        /**
         * The date data validation type.
         */
        date,

        /**
         * The time data validation type.
         */
        time,

        /**
         * The text length data validation type.
         */
        textLength,

        /**
         * The custom data validation type.
         */
        custom,

        /**
         * Inconsistent means that the range has inconsistent data validation, indicating that there are different rules on different cells.
         */
        inconsistent,

        /**
         * Mixed criteria means that the range has data validation present on some but not all cells.
         */
        mixedCriteria,
    }

    /**
     * Represents the data validation operator enum.
     */
    enum DataValidationOperator {
        between,

        notBetween,

        equalTo,

        notEqualTo,

        greaterThan,

        lessThan,

        greaterThanOrEqualTo,

        lessThanOrEqualTo,
    }

    /**
     * Represents the data validation error alert style. The default is `Stop`.
     */
    enum DataValidationAlertStyle {
        stop,

        warning,

        information,
    }

    enum DeleteShiftDirection {
        up,

        left,
    }

    enum DynamicFilterCriteria {
        unknown,

        aboveAverage,

        allDatesInPeriodApril,

        allDatesInPeriodAugust,

        allDatesInPeriodDecember,

        allDatesInPeriodFebruary,

        allDatesInPeriodJanuary,

        allDatesInPeriodJuly,

        allDatesInPeriodJune,

        allDatesInPeriodMarch,

        allDatesInPeriodMay,

        allDatesInPeriodNovember,

        allDatesInPeriodOctober,

        allDatesInPeriodQuarter1,

        allDatesInPeriodQuarter2,

        allDatesInPeriodQuarter3,

        allDatesInPeriodQuarter4,

        allDatesInPeriodSeptember,

        belowAverage,

        lastMonth,

        lastQuarter,

        lastWeek,

        lastYear,

        nextMonth,

        nextQuarter,

        nextWeek,

        nextYear,

        thisMonth,

        thisQuarter,

        thisWeek,

        thisYear,

        today,

        tomorrow,

        yearToDate,

        yesterday,
    }

    enum FilterDatetimeSpecificity {
        year,

        month,

        day,

        hour,

        minute,

        second,
    }

    enum FilterOn {
        bottomItems,

        bottomPercent,

        cellColor,

        dynamic,

        fontColor,

        values,

        topItems,

        topPercent,

        icon,

        custom,
    }

    enum FilterOperator {
        and,

        or,
    }

    enum HorizontalAlignment {
        general,

        left,

        center,

        right,

        fill,

        justify,

        centerAcrossSelection,

        distributed,
    }

    enum IconSet {
        invalid,

        threeArrows,

        threeArrowsGray,

        threeFlags,

        threeTrafficLights1,

        threeTrafficLights2,

        threeSigns,

        threeSymbols,

        threeSymbols2,

        fourArrows,

        fourArrowsGray,

        fourRedToBlack,

        fourRating,

        fourTrafficLights,

        fiveArrows,

        fiveArrowsGray,

        fiveRating,

        fiveQuarters,

        threeStars,

        threeTriangles,

        fiveBoxes,
    }

    enum ImageFittingMode {
        fit,

        fitAndCenter,

        fill,
    }

    enum InsertShiftDirection {
        down,

        right,
    }

    enum NamedItemScope {
        worksheet,

        workbook,
    }

    enum NamedItemType {
        string,

        integer,

        double,

        boolean,

        range,

        error,

        array,
    }

    enum RangeUnderlineStyle {
        none,

        single,

        double,

        singleAccountant,

        doubleAccountant,
    }

    enum SheetVisibility {
        visible,

        hidden,

        veryHidden,
    }

    enum RangeValueType {
        unknown,

        empty,

        string,

        integer,

        double,

        boolean,

        error,

        richValue,
    }

    enum KeyboardDirection {
        left,

        right,

        up,

        down,
    }

    /**
     * Specifies the search direction.
     */
    enum SearchDirection {
        /**
         * Search in forward order.
         */
        forward,

        /**
         * Search in reverse order.
         */
        backwards,
    }

    enum SortOrientation {
        rows,

        columns,
    }

    enum SortOn {
        value,

        cellColor,

        fontColor,

        icon,
    }

    enum SortDataOption {
        normal,

        textAsNumber,
    }

    enum SortMethod {
        pinYin,

        strokeCount,
    }

    enum VerticalAlignment {
        top,

        center,

        bottom,

        justify,

        distributed,
    }

    enum DocumentPropertyType {
        number,

        boolean,

        date,

        string,

        float,
    }

    enum SubtotalLocationType {
        /**
         * Subtotals are at the top.
         */
        atTop,

        /**
         * Subtotals are at the bottom.
         */
        atBottom,

        /**
         * Subtotals are off.
         */
        off,
    }

    enum PivotLayoutType {
        /**
         * A horizontally compressed form with labels from the next field in the same column.
         */
        compact,

        /**
         * Inner fields' items are always on a new line relative to the outer fields' items.
         */
        tabular,

        /**
         * Inner fields' items are on same row as outer fields' items and subtotals are always on the bottom.
         */
        outline,
    }

    enum ProtectionSelectionMode {
        /**
         * Selection is allowed for all cells.
         */
        normal,

        /**
         * Selection is allowed only for cells that are not locked.
         */
        unlocked,

        /**
         * Selection is not allowed for any cells.
         */
        none,
    }

    enum PageOrientation {
        portrait,

        landscape,
    }

    enum PaperType {
        letter,

        letterSmall,

        tabloid,

        ledger,

        legal,

        statement,

        executive,

        a3,

        a4,

        a4Small,

        a5,

        b4,

        b5,

        folio,

        quatro,

        paper10x14,

        paper11x17,

        note,

        envelope9,

        envelope10,

        envelope11,

        envelope12,

        envelope14,

        csheet,

        dsheet,

        esheet,

        envelopeDL,

        envelopeC5,

        envelopeC3,

        envelopeC4,

        envelopeC6,

        envelopeC65,

        envelopeB4,

        envelopeB5,

        envelopeB6,

        envelopeItaly,

        envelopeMonarch,

        envelopePersonal,

        fanfoldUS,

        fanfoldStdGerman,

        fanfoldLegalGerman,
    }

    enum ReadingOrder {
        /**
         * Reading order is determined by the language of the first character entered.
         * If a right-to-left language character is entered first, reading order is right to left.
         * If a left-to-right language character is entered first, reading order is left to right.
         */
        context,

        /**
         * Left to right reading order
         */
        leftToRight,

        /**
         * Right to left reading order
         */
        rightToLeft,
    }

    enum BuiltInStyle {
        normal,

        comma,

        currency,

        percent,

        wholeComma,

        wholeDollar,

        hlink,

        hlinkTrav,

        note,

        warningText,

        emphasis1,

        emphasis2,

        emphasis3,

        sheetTitle,

        heading1,

        heading2,

        heading3,

        heading4,

        input,

        output,

        calculation,

        checkCell,

        linkedCell,

        total,

        good,

        bad,

        neutral,

        accent1,

        accent1_20,

        accent1_40,

        accent1_60,

        accent2,

        accent2_20,

        accent2_40,

        accent2_60,

        accent3,

        accent3_20,

        accent3_40,

        accent3_60,

        accent4,

        accent4_20,

        accent4_40,

        accent4_60,

        accent5,

        accent5_20,

        accent5_40,

        accent5_60,

        accent6,

        accent6_20,

        accent6_40,

        accent6_60,

        explanatoryText,
    }

    enum PrintErrorType {
        asDisplayed,

        blank,

        dash,

        notAvailable,
    }

    enum WorksheetPositionType {
        none,

        before,

        after,

        beginning,

        end,
    }

    enum PrintComments {
        /**
         * Comments will not be printed.
         */
        noComments,

        /**
         * Comments will be printed as end notes at the end of the worksheet.
         */
        endSheet,

        /**
         * Comments will be printed where they were inserted in the worksheet.
         */
        inPlace,
    }

    enum PrintOrder {
        /**
         * Process down the rows before processing across pages or page fields to the right.
         */
        downThenOver,

        /**
         * Process across pages or page fields to the right before moving down the rows.
         */
        overThenDown,
    }

    enum PrintMarginUnit {
        /**
         * Assign the page margins in points. A point is 1/72 of an inch.
         */
        points,

        /**
         * Assign the page margins in inches.
         */
        inches,

        /**
         * Assign the page margins in centimeters.
         */
        centimeters,
    }

    enum HeaderFooterState {
        /**
         * Only one general header/footer is used for all pages printed.
         */
        default,

        /**
         * There is a separate first page header/footer, and a general header/footer used for all other pages.
         */
        firstAndDefault,

        /**
         * There is a different header/footer for odd and even pages.
         */
        oddAndEven,

        /**
         * There is a separate first page header/footer, then there is a separate header/footer for odd and even pages.
         */
        firstOddAndEven,
    }

    /**
     * The behavior types when AutoFill is used on a range in the workbook.
     */
    enum AutoFillType {
        /**
         * Populates the adjacent cells based on the surrounding data (the standard AutoFill behavior).
         */
        fillDefault,

        /**
         * Populates the adjacent cells with data based on the selected data.
         */
        fillCopy,

        /**
         * Populates the adjacent cells with data that follows a pattern in the copied cells.
         */
        fillSeries,

        /**
         * Populates the adjacent cells with the selected formulas.
         */
        fillFormats,

        /**
         * Populates the adjacent cells with the selected values.
         */
        fillValues,

        /**
         * A version of "FillSeries" for dates that bases the pattern on either the day of the month or the day of the week, depending on the context.
         */
        fillDays,

        /**
         * A version of "FillSeries" for dates that bases the pattern on the day of the week and only includes weekdays.
         */
        fillWeekdays,

        /**
         * A version of "FillSeries" for dates that bases the pattern on the month.
         */
        fillMonths,

        /**
         * A version of "FillSeries" for dates that bases the pattern on the year.
         */
        fillYears,

        /**
         * A version of "FillSeries" for numbers that fills out the values in the adjacent cells according to a linear trend model.
         */
        linearTrend,

        /**
         * A version of "FillSeries" for numbers that fills out the values in the adjacent cells according to a growth trend model.
         */
        growthTrend,

        /**
         * Populates the adjacent cells by using Excel's Flash Fill feature.
         */
        flashFill,
    }

    enum GroupOption {
        /**
         * Group by rows.
         */
        byRows,

        /**
         * Group by columns.
         */
        byColumns,
    }

    enum RangeCopyType {
        all,

        formulas,

        values,

        formats,

        link,
    }

    enum LinkedDataTypeState {
        none,

        validLinkedData,

        disambiguationNeeded,

        brokenLinkedData,

        fetchingData,
    }

    /**
     * Specifies the shape type for a `GeometricShape` object.
     */
    enum GeometricShapeType {
        lineInverse,

        triangle,

        rightTriangle,

        rectangle,

        diamond,

        parallelogram,

        trapezoid,

        nonIsoscelesTrapezoid,

        pentagon,

        hexagon,

        heptagon,

        octagon,

        decagon,

        dodecagon,

        star4,

        star5,

        star6,

        star7,

        star8,

        star10,

        star12,

        star16,

        star24,

        star32,

        roundRectangle,

        round1Rectangle,

        round2SameRectangle,

        round2DiagonalRectangle,

        snipRoundRectangle,

        snip1Rectangle,

        snip2SameRectangle,

        snip2DiagonalRectangle,

        plaque,

        ellipse,

        teardrop,

        homePlate,

        chevron,

        pieWedge,

        pie,

        blockArc,

        donut,

        noSmoking,

        rightArrow,

        leftArrow,

        upArrow,

        downArrow,

        stripedRightArrow,

        notchedRightArrow,

        bentUpArrow,

        leftRightArrow,

        upDownArrow,

        leftUpArrow,

        leftRightUpArrow,

        quadArrow,

        leftArrowCallout,

        rightArrowCallout,

        upArrowCallout,

        downArrowCallout,

        leftRightArrowCallout,

        upDownArrowCallout,

        quadArrowCallout,

        bentArrow,

        uturnArrow,

        circularArrow,

        leftCircularArrow,

        leftRightCircularArrow,

        curvedRightArrow,

        curvedLeftArrow,

        curvedUpArrow,

        curvedDownArrow,

        swooshArrow,

        cube,

        can,

        lightningBolt,

        heart,

        sun,

        moon,

        smileyFace,

        irregularSeal1,

        irregularSeal2,

        foldedCorner,

        bevel,

        frame,

        halfFrame,

        corner,

        diagonalStripe,

        chord,

        arc,

        leftBracket,

        rightBracket,

        leftBrace,

        rightBrace,

        bracketPair,

        bracePair,

        callout1,

        callout2,

        callout3,

        accentCallout1,

        accentCallout2,

        accentCallout3,

        borderCallout1,

        borderCallout2,

        borderCallout3,

        accentBorderCallout1,

        accentBorderCallout2,

        accentBorderCallout3,

        wedgeRectCallout,

        wedgeRRectCallout,

        wedgeEllipseCallout,

        cloudCallout,

        cloud,

        ribbon,

        ribbon2,

        ellipseRibbon,

        ellipseRibbon2,

        leftRightRibbon,

        verticalScroll,

        horizontalScroll,

        wave,

        doubleWave,

        plus,

        flowChartProcess,

        flowChartDecision,

        flowChartInputOutput,

        flowChartPredefinedProcess,

        flowChartInternalStorage,

        flowChartDocument,

        flowChartMultidocument,

        flowChartTerminator,

        flowChartPreparation,

        flowChartManualInput,

        flowChartManualOperation,

        flowChartConnector,

        flowChartPunchedCard,

        flowChartPunchedTape,

        flowChartSummingJunction,

        flowChartOr,

        flowChartCollate,

        flowChartSort,

        flowChartExtract,

        flowChartMerge,

        flowChartOfflineStorage,

        flowChartOnlineStorage,

        flowChartMagneticTape,

        flowChartMagneticDisk,

        flowChartMagneticDrum,

        flowChartDisplay,

        flowChartDelay,

        flowChartAlternateProcess,

        flowChartOffpageConnector,

        actionButtonBlank,

        actionButtonHome,

        actionButtonHelp,

        actionButtonInformation,

        actionButtonForwardNext,

        actionButtonBackPrevious,

        actionButtonEnd,

        actionButtonBeginning,

        actionButtonReturn,

        actionButtonDocument,

        actionButtonSound,

        actionButtonMovie,

        gear6,

        gear9,

        funnel,

        mathPlus,

        mathMinus,

        mathMultiply,

        mathDivide,

        mathEqual,

        mathNotEqual,

        cornerTabs,

        squareTabs,

        plaqueTabs,

        chartX,

        chartStar,

        chartPlus,
    }

    enum ConnectorType {
        straight,

        elbow,

        curve,
    }

    enum ContentType {
        /**
         * Indicates a plain format type for the comment content.
         */
        plain,

        /**
         * Comment content containing mentions.
         */
        mention,
    }

    enum SpecialCellType {
        /**
         * All cells with conditional formats.
         */
        conditionalFormats,

        /**
         * Cells with validation criteria.
         */
        dataValidations,

        /**
         * Cells with no content.
         */
        blanks,

        /**
         * Cells containing constants.
         */
        constants,

        /**
         * Cells containing formulas.
         */
        formulas,

        /**
         * Cells with the same conditional format as the first cell in the range.
         */
        sameConditionalFormat,

        /**
         * Cells with the same data validation criteria as the first cell in the range.
         */
        sameDataValidation,

        /**
         * Cells that are visible.
         */
        visible,
    }

    enum SpecialCellValueType {
        /**
         * Cells that have errors, boolean, numeric, or string values.
         */
        all,

        /**
         * Cells that have errors.
         */
        errors,

        /**
         * Cells that have errors or boolean values.
         */
        errorsLogical,

        /**
         * Cells that have errors or numeric values.
         */
        errorsNumbers,

        /**
         * Cells that have errors or string values.
         */
        errorsText,

        /**
         * Cells that have errors, boolean, or numeric values.
         */
        errorsLogicalNumber,

        /**
         * Cells that have errors, boolean, or string values.
         */
        errorsLogicalText,

        /**
         * Cells that have errors, numeric, or string values.
         */
        errorsNumberText,

        /**
         * Cells that have a boolean value.
         */
        logical,

        /**
         * Cells that have a boolean or numeric value.
         */
        logicalNumbers,

        /**
         * Cells that have a boolean or string value.
         */
        logicalText,

        /**
         * Cells that have a boolean, numeric, or string value.
         */
        logicalNumbersText,

        /**
         * Cells that have a numeric value.
         */
        numbers,

        /**
         * Cells that have a numeric or string value.
         */
        numbersText,

        /**
         * Cells that have a string value.
         */
        text,
    }

    /**
     * Specifies the way that an object is attached to its underlying cells.
     */
    enum Placement {
        /**
         * The object is moved and sized with the cells.
         */
        twoCell,

        /**
         * The object is moved with the cells.
         */
        oneCell,

        /**
         * The object is free floating.
         */
        absolute,
    }

    enum FillPattern {
        none,

        solid,

        gray50,

        gray75,

        gray25,

        horizontal,

        vertical,

        down,

        up,

        checker,

        semiGray75,

        lightHorizontal,

        lightVertical,

        lightDown,

        lightUp,

        grid,

        crissCross,

        gray16,

        gray8,

        linearGradient,

        rectangularGradient,
    }

    /**
     * Specifies the horizontal alignment for the text frame in a shape.
     */
    enum ShapeTextHorizontalAlignment {
        left,

        center,

        right,

        justify,

        justifyLow,

        distributed,

        thaiDistributed,
    }

    /**
     * Specifies the vertical alignment for the text frame in a shape.
     */
    enum ShapeTextVerticalAlignment {
        top,

        middle,

        bottom,

        justified,

        distributed,
    }

    /**
     * Specifies the vertical overflow for the text frame in a shape.
     */
    enum ShapeTextVerticalOverflow {
        /**
         * Allow text to overflow the text frame vertically (can be from the top, bottom, or both depending on the text alignment).
         */
        overflow,

        /**
         * Hide text that does not fit vertically within the text frame, and add an ellipsis (...) at the end of the visible text.
         */
        ellipsis,

        /**
         * Hide text that does not fit vertically within the text frame.
         */
        clip,
    }

    /**
     * Specifies the horizontal overflow for the text frame in a shape.
     */
    enum ShapeTextHorizontalOverflow {
        overflow,

        clip,
    }

    /**
     * Specifies the reading order for the text frame in a shape.
     */
    enum ShapeTextReadingOrder {
        leftToRight,

        rightToLeft,
    }

    /**
     * Specifies the orientation for the text frame in a shape.
     */
    enum ShapeTextOrientation {
        horizontal,

        vertical,

        vertical270,

        wordArtVertical,

        eastAsianVertical,

        mongolianVertical,

        wordArtVerticalRTL,
    }

    /**
     * Determines the type of automatic sizing allowed.
     */
    enum ShapeAutoSize {
        /**
         * No autosizing.
         */
        autoSizeNone,

        /**
         * The text is adjusted to fit the shape.
         */
        autoSizeTextToFitShape,

        /**
         * The shape is adjusted to fit the text.
         */
        autoSizeShapeToFitText,

        /**
         * A combination of automatic sizing schemes are used.
         */
        autoSizeMixed,
    }

    /**
     * Specifies the slicer sort behavior for `Slicer.sortBy`.
     */
    enum SlicerSortType {
        /**
         * Sort slicer items in the order provided by the data source.
         */
        dataSourceOrder,

        /**
         * Sort slicer items in ascending order by item captions.
         */
        ascending,

        /**
         * Sort slicer items in descending order by item captions.
         */
        descending,
    }

    /**
     * Represents a category of number formats.
     */
    enum NumberFormatCategory {
        /**
         * General format cells have no specific number format.
         */
        general,

        /**
         * Number is used for general display of numbers. Currency and Accounting offer specialized formatting for monetary value.
         */
        number,

        /**
         * Currency formats are used for general monetary values. Use Accounting formats to align decimal points in a column.
         */
        currency,

        /**
         * Accounting formats line up the currency symbols and decimal points in a column.
         */
        accounting,

        /**
         * Date formats display date and time serial numbers as date values. Date formats that begin with an asterisk (*) respond to changes in regional date and time settings that are specified for the operating system. Formats without an asterisk are not affected by operating system settings.
         */
        date,

        /**
         * Time formats display date and time serial numbers as date values. Time formats that begin with an asterisk (*) respond to changes in regional date and time settings that are specified for the operating system. Formats without an asterisk are not affected by operating system settings.
         */
        time,

        /**
         * Percentage formats multiply the cell value by 100 and displays the result with a percent symbol.
         */
        percentage,

        /**
         * Fraction formats display the cell value as a whole number with the remainder rounded to the nearest fraction value.
         */
        fraction,

        /**
         * Scientific formats display the cell value as a number between 1 and 10 multiplied by a power of 10.
         */
        scientific,

        /**
         * Text format cells are treated as text even when a number is in the cell. The cell is displayed exactly as entered.
         */
        text,

        /**
         * Special formats are useful for tracking list and database values.
         */
        special,

        /**
         * A custom format that is not a part of any category.
         */
        custom,
    }

    //
    // Type
    //

    type RangeValue = string | number | boolean | null | undefined;
}
