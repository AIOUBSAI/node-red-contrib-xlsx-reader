const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

module.exports = function(RED) {
    function XLSXReaderNode(config) {
        RED.nodes.createNode(this, config);
        const node = this;

        // Config
        node.path = config.path;
        node.pathType = config.pathType || "str";  // str | msg | flow | global | env
        node.mode = config.mode || "file";         // file | directory
        node.excludeRegex = config.excludeRegex || "";
        node.includeHidden = !!config.includeHidden;

        // New config
        node.outputTargetType = config.outputTargetType || "msg";   // msg | flow | global
        node.outputTargetPath = config.outputTargetPath || "data";  // deep path under chosen scope
        node.fillMerged       = !!config.fillMerged;                // forward-fill columns
        node.mergedColumnsRaw = config.mergedColumns || "";         // comma-separated headers (case-sensitive)

        node.on("input", async function(msg, send, done) {
            try {
                node.status({ fill: "blue", shape: "dot", text: "reading..." });

                // Resolve path (supports str, msg, flow, global, env)
                let filePath = RED.util.evaluateNodeProperty(node.path, node.pathType, node, msg);
                if (!filePath) throw new Error("No path specified.");

                // List files
                let files = [];
                if (node.mode === "directory") {
                    files = fs.readdirSync(filePath)
                        .filter(f => f.toLowerCase().endsWith(".xlsx"))
                        .map(f => path.join(filePath, f))
                        .filter(f => {
                            try { return fs.statSync(f).isFile(); }
                            catch { return false; }
                        });
                } else {
                    files = [filePath];
                }
                if (!files.length) throw new Error("No .xlsx files found for the provided path/mode.");

                // Aggregator
                const resultMap = {}; // { [filename]: { [sheet]: rows[] } }
                let fileCount = 0, sheetCount = 0, rowCount = 0;

                const exclude = node.excludeRegex ? new RegExp(node.excludeRegex) : null;
                const mergedColumns = node.mergedColumnsRaw
                    .split(",")
                    .map(s => s.trim())
                    .filter(Boolean); // case-sensitive

                // Read all files
                for (const f of files) {
                    const workbook = XLSX.readFile(f, { cellDates: true });
                    const sheets = workbook.SheetNames;
                    resultMap[f] = {};
                    fileCount++;

                    for (const sheetName of sheets) {
                        const ws = workbook.Sheets[sheetName];

                        // Hidden/veryHidden skip
                        const isHidden = workbook.Workbook &&
                                         Array.isArray(workbook.Workbook.Sheets) &&
                                         !!workbook.Workbook.Sheets.find(s => s.name === sheetName && s.Hidden);
                        if (!node.includeHidden && isHidden) continue;

                        // Exclude regex on sheet name
                        if (exclude && exclude.test(sheetName)) continue;

                        // Convert to JSON
                        //const data = XLSX.utils.sheet_to_json(ws, { defval: null });
                        let data = XLSX.utils.sheet_to_json(ws, { defval: null, blankrows: false });

                            // Remove XLSX placeholder columns like "__EMPTY", "__EMPTY_1", etc.
                            data = data.map(row => {
                            for (const key of Object.keys(row)) {
                                if (key.startsWith("__EMPTY")) {
                                delete row[key];
                                }
                            }
                            return row;
                            });



                        // Optional forward-fill for given columns
                        if (node.fillMerged && mergedColumns.length) {
                            fillForward(data, mergedColumns);
                        }

                        resultMap[f][sheetName] = data;
                        sheetCount++;
                        rowCount += Array.isArray(data) ? data.length : 0;
                    }
                }

                // Build aggregated payload
                const aggregated = {
                    data: resultMap,
                    summary: { fileCount, sheetCount, rowCount },
                    options: {
                        excludeRegex: node.excludeRegex || null,
                        includeHidden: !!node.includeHidden,
                        fillMerged: !!node.fillMerged,
                        mergedColumns
                    }
                };

                // Place output in chosen scope/path
                setOutput(RED, node, msg, aggregated);

                // Emit a single message
                node.status({ fill: "green", shape: "dot", text: `files:${fileCount} sheets:${sheetCount}` });
                send(msg);
                if (done) done();
            } catch (err) {
                node.error(err, msg);
                node.status({ fill: "red", shape: "ring", text: "error" });
                if (done) done(err);
            }
        });
    }

    RED.nodes.registerType("xlsx-reader", XLSXReaderNode);

    // --- helpers ---

    function isEmpty(v) {
        return v === null || v === undefined || (typeof v === "string" && v.trim() === "");
    }

    /**
     * Forward-fill empty values in specified columns (by header names).
     * This is a pragmatic approach to handle Excel merged-cells exported via sheet_to_json.
     * Case-sensitive header names are expected.
     */
    function fillForward(rows, columns) {
        if (!Array.isArray(rows) || !rows.length) return;
        for (const col of columns) {
            let last = undefined;
            for (let i = 0; i < rows.length; i++) {
                const v = rows[i][col];
                if (isEmpty(v)) {
                    if (!isEmpty(last)) rows[i][col] = last;
                } else {
                    last = v;
                }
            }
        }
    }

    /**
     * Set aggregated output into msg/flow/global at a deep path.
     * - msg: uses RED.util.setMessageProperty (creates path as needed).
     * - flow/global: supports deep path `root.child.leaf`.
     */
    function setOutput(RED, node, msg, value) {
        const scope = node.outputTargetType || "msg";
        const pathStr = node.outputTargetPath || "data";

        if (scope === "msg") {
            RED.util.setMessageProperty(msg, pathStr, value, true);
            return;
        }

        // flow/global store
        const ctx = node.context()[scope];
        if (!ctx || typeof ctx.get !== "function" || typeof ctx.set !== "function") {
            // Fallback: if something odd, write to msg
            RED.util.setMessageProperty(msg, pathStr, value, true);
            return;
        }

        if (!pathStr || !pathStr.trim()) {
            // No deep path → store in a default key
            ctx.set("xlsxReader", value);
            return;
        }

        const parts = pathStr.split(".").filter(Boolean);
        const rootKey = parts.shift();
        if (!rootKey) {
            ctx.set("xlsxReader", value);
            return;
        }

        if (parts.length === 0) {
            // Direct root set
            ctx.set(rootKey, value);
            return;
        }

        // Deep set under rootKey
        let rootObj = ctx.get(rootKey);
        if (typeof rootObj !== "object" || rootObj === null) rootObj = {};
        setDeep(rootObj, parts, value);
        ctx.set(rootKey, rootObj);
    }

    function setDeep(obj, pathArr, val) {
        let cur = obj;
        for (let i = 0; i < pathArr.length - 1; i++) {
            const k = pathArr[i];
            if (typeof cur[k] !== "object" || cur[k] === null) cur[k] = {};
            cur = cur[k];
        }
        cur[pathArr[pathArr.length - 1]] = val;
    }
};
