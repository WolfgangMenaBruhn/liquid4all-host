window.excelInterop = {
    isReady: false,
    host: null,
    skipped: false,

    isOfficeEnvironment() {
        if (window.location.href.includes("_host_Info")) {
            return true;
        }

        try {
            if (window.sessionStorage.getItem("hostInfoValue")) {
                return true;
            }
        } catch {
            // sessionStorage can throw in a sandboxed frame
        }

        return false;
    },

    start() {
        this._nativePushState = window.history.pushState.bind(window.history);
        this._nativeReplaceState = window.history.replaceState.bind(window.history);

        if (!this.isOfficeEnvironment()) {
            this.skipped = true;
            this.isReady = false;
            this.host = null;
            return;
        }

        const officeBase = "https://appsforoffice.microsoft.com/lib/1/hosted/";
        const locale = this.getHostLocale();
        const afterOfficeJs = () => {
            try {
                Office.initialize = Office.initialize || function () { };
            } catch {
                // Office may still be constructing its namespace.
            }

            Office.onReady((info) => {
                this.host = info.host ? String(info.host) : null;
                this.isReady = true;
                this.applyTaskPaneWidth();
                this._startBlazorIfPending();
            });
        };

        const loadOfficeJs = () => this._loadScript(officeBase + "office.js")
            .then(afterOfficeJs)
            .catch(() => {
                this.isReady = false;
                this.host = null;
                this._startBlazorIfPending();
            });

        // English strings are bundled in office.js. Other Excel UI languages fetch
        // {locale}/office_strings.js first; if that file is missing, onReady never
        // fires and Excel shows the "check your network" add-in error.
        if (locale && locale !== "en-us") {
            this._loadScript(officeBase + locale + "/office_strings.js")
                .catch(() => {})
                .then(loadOfficeJs);
            return;
        }

        loadOfficeJs();
    },

    startBlazor() {
        const begin = () => {
            if (this._blazorStarted || typeof Blazor === "undefined") {
                return;
            }

            this._blazorStarted = true;
            // Excel's WebView reports the Office display language as navigator.language.
            // Starting Blazor in that culture downloads _framework/de/*.wasm before the
            // app can render. A failed or slow satellite fetch looks like a network error
            // in Excel. Boot in English, then load German resources after startup.
            Promise.resolve(Blazor.start({ applicationCulture: "en-US" })).catch((error) => {
                console.error("Liquid4All: Blazor.start failed", error);
            });
        };

        if (this.skipped || this.isReady) {
            begin();
            return;
        }

        this._pendingBlazorStart = begin;
        window.setTimeout(() => this._startBlazorIfPending(), 15000);
    },

    _startBlazorIfPending() {
        if (typeof this._pendingBlazorStart === "function") {
            const begin = this._pendingBlazorStart;
            this._pendingBlazorStart = null;
            begin();
        }
    },

    getHostLocale() {
        try {
            const query = new URLSearchParams(window.location.search);
            let hostInfo = query.get("_host_Info") || query.get("_host_info") || "";
            if (!hostInfo) {
                hostInfo = window.sessionStorage.getItem("hostInfoValue") || "";
            }

            hostInfo = decodeURIComponent(hostInfo);
            const parts = hostInfo.includes("$") ? hostInfo.split("$") : hostInfo.split("|");
            const locale = String(parts[3] || "").toLowerCase();
            if (!locale) {
                return "";
            }

            if (locale.startsWith("de")) {
                return "de-de";
            }

            return locale;
        } catch {
            return "";
        }
    },

    getDisplayLanguage() {
        try {
            if (typeof Office !== "undefined" && Office.context && Office.context.displayLanguage) {
                return String(Office.context.displayLanguage);
            }
        } catch {
            // Office.context can throw before onReady
        }

        return "";
    },

    getExcelCulture() {
        const language = (this.getDisplayLanguage() || this.getHostLocale()).toLowerCase();
        return language.startsWith("de") ? "de" : "en";
    },

    applyExcelCulture() {
        const culture = this.getExcelCulture();
        document.documentElement.lang = culture;
        return culture;
    },

    async loadSatelliteCultures(cultures) {
        const list = Array.isArray(cultures) ? cultures : [cultures];
        const loaders = [
            globalThis.INTERNAL,
            typeof Blazor !== "undefined" ? Blazor.runtime && Blazor.runtime.INTERNAL : null,
            typeof Blazor !== "undefined" ? Blazor._internal : null
        ];

        for (const api of loaders) {
            if (api && typeof api.loadSatelliteAssemblies === "function") {
                try {
                    await api.loadSatelliteAssemblies(list);
                    return true;
                } catch (error) {
                    console.warn("Liquid4All: satellite assembly load failed", error);
                }
            }
        }

        return false;
    },

    _loadScript(src) {
        return new Promise((resolve, reject) => {
            const script = document.createElement("script");
            script.src = src;
            script.async = false;
            script.onload = () => resolve();
            script.onerror = () => reject(new Error("Failed to load " + src));
            document.head.appendChild(script);
        });
    },

    applyTaskPaneWidth() {
        try {
            if (typeof Office === "undefined"
                || !Office.context
                || !Office.context.requirements.isSetSupported("TaskPaneApi", "1.1")) {
                return;
            }

            Office.extensionLifeCycle.taskpane.setWidth(700);
        } catch {
            // Older Excel hosts throw or ignore when the task pane API is missing.
        }
    },

    navigate(path) {
        const url = new URL(window.location.href);
        const baseHref = document.querySelector("base")?.getAttribute("href") || "/";
        const baseUrl = new URL(baseHref, url.origin);
        const relative = !path || path === "/" ? "./" : String(path).replace(/^\//, "");
        const nextUrl = new URL(relative, baseUrl);
        nextUrl.search = url.search;
        nextUrl.hash = "";
        const next = `${nextUrl.pathname}${nextUrl.search}`;
        const push = this._nativePushState || window.history.pushState.bind(window.history);
        push(null, "", next);
        window.dispatchEvent(new PopStateEvent("popstate"));
    },

    getStatus() {
        return {
            ok: this.isReady && this.host === "Excel",
            host: this.host,
            skipped: this.skipped
        };
    },

    async waitForHost(timeoutMs = 8000) {
        if (this.skipped) {
            return null;
        }

        const start = Date.now();
        while (!this.isReady && Date.now() - start < timeoutMs) {
            await new Promise((resolve) => setTimeout(resolve, 50));
        }

        return this.host;
    },

    async whenReady(timeoutMs = 8000) {
        await this.waitForHost(timeoutMs);
        return this.getStatus();
    },

    async readUsedRange(sheetName) {
        return await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
            sheet.load("isNullObject");
            await context.sync();
            if (sheet.isNullObject) {
                return null;
            }

            const range = sheet.getUsedRangeOrNullObject();
            range.load(["values", "rowCount", "columnCount"]);
            await context.sync();
            if (range.isNullObject) {
                return { values: [], rowCount: 0, columnCount: 0 };
            }

            return {
                values: range.values.map(row => row.map(cell => cell == null ? null : String(cell))),
                rowCount: range.rowCount,
                columnCount: range.columnCount
            };
        });
    },

    async writeSheet(sheetName, values, showSheet) {
        await Excel.run(async (context) => {
            let sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
            sheet.load("isNullObject");
            await context.sync();

            if (sheet.isNullObject) {
                sheet = context.workbook.worksheets.add(sheetName);
            }

            sheet.visibility = showSheet
                ? Excel.SheetVisibility.visible
                : Excel.SheetVisibility.veryHidden;

            const used = sheet.getUsedRangeOrNullObject();
            used.load("isNullObject");
            await context.sync();
            if (!used.isNullObject) {
                used.clear();
            }

            const rowCount = values.length;
            const columnCount = rowCount === 0 ? 1 : values[0].length;
            const target = sheet.getRangeByIndexes(0, 0, Math.max(rowCount, 1), Math.max(columnCount, 1));
            target.numberFormat = "@";
            if (rowCount > 0) {
                target.values = values;
                sheet.getRangeByIndexes(0, 0, 1, columnCount).format.font.bold = true;
            }

            await context.sync();
        });
    },

    async setSheetVisibility(sheetName, showSheet) {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
            sheet.load("isNullObject");
            await context.sync();
            if (sheet.isNullObject) {
                return;
            }

            sheet.visibility = showSheet
                ? Excel.SheetVisibility.visible
                : Excel.SheetVisibility.veryHidden;
            await context.sync();
        });
    },

    async exportCalculation(sheetName, headers, rows, chartOptions, chartRows) {
        const pivotStartColumn = 26;
        const pivotName = "L4A_Pivot_" + Date.now();
        const chartSource = this._chartSourceValues(headers, rows, chartRows);

        await Excel.run(async (context) => {
            const existing = context.workbook.worksheets.getItemOrNullObject(sheetName);
            existing.load("isNullObject");
            await context.sync();
            if (!existing.isNullObject) {
                existing.delete();
                await context.sync();
            }

            const sheet = context.workbook.worksheets.add(sheetName);
            const columnCount = headers.length;
            const rowCount = 1 + rows.length;
            const tableValues = [headers, ...rows.map((row) => {
                const next = [...row];
                next[0] = this._toExcelSerialDate(row[0]);
                next[4] = Number(row[4]);
                next[5] = Number(row[5]);
                return next;
            })];
            const range = sheet.getRangeByIndexes(0, 0, rowCount, columnCount);
            range.values = tableValues;

            const headerRange = sheet.getRangeByIndexes(0, 0, 1, columnCount);
            headerRange.format.font.bold = true;
            headerRange.format.fill.color = "#D3D3D3";
            headerRange.format.horizontalAlignment = Excel.HorizontalAlignment.center;
            this._applyThinBorders(headerRange, false);

            const dataBody = sheet.getRangeByIndexes(1, 0, rows.length, columnCount);
            this._applyThinBorders(dataBody, rows.length > 1 && columnCount > 1);

            sheet.getRangeByIndexes(1, 0, rows.length, 1).numberFormat = chartOptions?.dateFormat || "yyyy-mm-dd";
            sheet.getRangeByIndexes(1, 4, rows.length, 2).numberFormat = "0.00;[Red]-0.00";
            range.format.autofitColumns();

            const chartColumnCount = chartSource[0].length;
            const chartRange = sheet.getRangeByIndexes(
                0,
                pivotStartColumn,
                chartSource.length,
                chartColumnCount);
            chartRange.values = chartSource;
            sheet.getRangeByIndexes(1, pivotStartColumn, Math.max(chartSource.length - 1, 1), 1)
                .numberFormat = chartOptions?.dateFormat || "yyyy-mm-dd";
            sheet.getRangeByIndexes(1, pivotStartColumn + 2, Math.max(chartSource.length - 1, 1), 1)
                .numberFormat = "0.00;[Red]-0.00";

            sheet.activate();
            await context.sync();

            const destination = sheet.getRangeByIndexes(0, pivotStartColumn + chartColumnCount + 1, 1, 1);
            sheet.pivotTables.add(pivotName, chartRange, destination);
            await context.sync();
        });

        try {
            await Excel.run(async (context) => {
                const pivotTable = context.workbook.worksheets.getItem(sheetName).pivotTables.getItem(pivotName);
                await this._configureCalculationPivotFields(pivotTable, headers, context);
            });
        } catch (error) {
            console.error("Liquid4All: could not configure calculation pivot fields", error);
        }

        const chartAdded = await this._addCalculationPivotChart(sheetName, pivotName, pivotStartColumn);
        if (chartAdded) {
            await this._formatCalculationChart(sheetName, chartOptions);
        }
    },

    _applyThinBorders(range, includeInside) {
        const edges = [
            Excel.BorderIndex.edgeTop,
            Excel.BorderIndex.edgeBottom,
            Excel.BorderIndex.edgeLeft,
            Excel.BorderIndex.edgeRight
        ];
        if (includeInside) {
            edges.push(Excel.BorderIndex.insideHorizontal, Excel.BorderIndex.insideVertical);
        }

        edges.forEach((edge) => {
            range.format.borders.getItem(edge).style = Excel.BorderLineStyle.continuous;
        });
    },

    _toExcelSerialDate(value) {
        if (typeof value === "number" && Number.isFinite(value)) {
            return value;
        }

        const text = String(value ?? "");
        const match = /^(\d{4})-(\d{2})-(\d{2})/.exec(text);
        if (!match) {
            return value;
        }

        const year = Number(match[1]);
        const month = Number(match[2]);
        const day = Number(match[3]);
        const utc = Date.UTC(year, month - 1, day);
        return (utc - Date.UTC(1899, 11, 30)) / 86400000;
    },

    _chartSourceValues(headers, rows, chartRows) {
        const header = [headers[0], headers[1], headers[5]];
        const sourceRows = Array.isArray(chartRows) && chartRows.length > 0
            ? chartRows.map((row) => [
                this._toExcelSerialDate(row[0]),
                row[1],
                Number(row[2])
            ])
            : this._lastBalanceRows(rows);

        return [header, ...sourceRows];
    },

    _lastBalanceRows(rows) {
        const last = new Map();
        for (const row of rows) {
            last.set(String(row[0]) + "\0" + String(row[1]), [
                this._toExcelSerialDate(row[0]),
                row[1],
                Number(row[5])
            ]);
        }

        return [...last.values()];
    },

    async _addCalculationPivotChart(sheetName, pivotName, pivotStartColumn) {
        const attempts = [
            async (context, sheet, pivotTable) =>
                sheet.charts.add(Excel.ChartType.line, pivotTable.layout.getRange()),
            async (context, sheet, pivotTable) =>
                sheet.charts.add(Excel.ChartType.line, pivotTable.layout.getDataBodyRange()),
            async (context, sheet, pivotTable) =>
                sheet.charts.add(Excel.ChartType.line, sheet.getRangeByIndexes(0, pivotStartColumn, 1, 1)),
            async (context, sheet, pivotTable) =>
                sheet.charts.add("Line", pivotTable.layout.getRange())
        ];

        for (const attempt of attempts) {
            try {
                await Excel.run(async (context) => {
                    const sheet = context.workbook.worksheets.getItem(sheetName);
                    const pivotTable = sheet.pivotTables.getItem(pivotName);
                    sheet.activate();
                    const chart = await attempt(context, sheet, pivotTable);
                    await context.sync();
                    sheet.getRangeByIndexes(0, pivotStartColumn, 1, 16).columnHidden = true;
                    await context.sync();
                    return chart;
                });
                return true;
            } catch (error) {
                console.error("Liquid4All: chart add attempt failed", error);
            }
        }

        return false;
    },

    async _formatCalculationChart(sheetName, chartOptions) {
        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getItem(sheetName);
                const chart = sheet.charts.getItemAt(0);
                chart.plotVisibleOnly = false;
                if (chartOptions?.title) {
                    chart.title.text = chartOptions.title;
                }
                if (Excel.ChartLegendPosition && Excel.ChartLegendPosition.bottom) {
                    chart.legend.position = Excel.ChartLegendPosition.bottom;
                }
                if (chartOptions?.dateAxisTitle) {
                    chart.axes.categoryAxis.hasTitle = true;
                    chart.axes.categoryAxis.title.text = chartOptions.dateAxisTitle;
                }
                if (chartOptions?.valueAxisTitle) {
                    chart.axes.valueAxis.hasTitle = true;
                    chart.axes.valueAxis.title.text = chartOptions.valueAxisTitle;
                }
                chart.axes.valueAxis.numberFormat = "#,##0.00;[Red]-#,##0.00";
                if (chartOptions?.dateFormat) {
                    chart.axes.categoryAxis.numberFormat = chartOptions.dateFormat;
                }
                chart.left = 480;
                chart.top = 0;
                chart.width = 520;
                chart.height = 400;
                this._showPivotChartFieldButtons(chart);
                await context.sync();

                await this._stylePivotChartSeries(chart, context);
            });
        } catch (error) {
            console.error("Liquid4All: could not format calculation chart", error);
        }
    },

    async _configureCalculationPivotFields(pivotTable, headers, context) {
        pivotTable.hierarchies.load("items/name");
        await context.sync();

        const names = (pivotTable.hierarchies.items || []).map((item) => item.name);
        const dateName = this._matchHierarchyName(names, headers[0]);
        const accountName = this._matchHierarchyName(names, headers[1]);
        const balanceName = this._matchHierarchyName(names, headers[5]);

        pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem(dateName));
        await context.sync();

        pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem(accountName));
        await context.sync();

        pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem(balanceName));
        await context.sync();

        try {
            const balanceField = pivotTable.dataHierarchies.getItemAt(0);
            balanceField.summarizeBy = (Excel.AggregationFunction && Excel.AggregationFunction.sum) || "Sum";
            pivotTable.layout.showRowGrandTotals = false;
            pivotTable.layout.showColumnGrandTotals = false;
            await context.sync();
        } catch (error) {
            console.error("Liquid4All: could not set pivot aggregation", error);
        }

        try {
            pivotTable.dataHierarchies.getItemAt(0).numberFormat = "0.00";
            await context.sync();
        } catch (error) {
            console.error("Liquid4All: could not set pivot number format", error);
        }
    },

    _matchHierarchyName(hierarchyNames, wanted) {
        if (wanted && hierarchyNames.includes(wanted)) {
            return wanted;
        }

        const wantedLower = String(wanted || "").toLowerCase();
        const match = hierarchyNames.find((name) => String(name).toLowerCase() === wantedLower);
        if (match) {
            return match;
        }

        throw new Error("Pivot field not found: " + wanted + " in [" + hierarchyNames.join(", ") + "]");
    },

    _showPivotChartFieldButtons(chart) {
        if (typeof Office === "undefined" || !Office.context.requirements.isSetSupported("ExcelApi", "1.9")) {
            return;
        }

        chart.pivotOptions.showAxisFieldButtons = true;
        chart.pivotOptions.showLegendFieldButtons = true;
        chart.pivotOptions.showValueFieldButtons = true;
        chart.pivotOptions.showReportFilterFieldButtons = false;
    },

    async _stylePivotChartSeries(chart, context) {
        chart.series.load("count");
        await context.sync();
        for (let i = 0; i < chart.series.count; i++) {
            const series = chart.series.getItemAt(i);
            series.markerStyle = "Circle";
            series.markerSize = 5;
        }
        await context.sync();
    }
};

window.excelInterop.start();
