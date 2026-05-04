/*
  Illustrator EPS batch text/layer probe (large-run friendly).

  What is new:
    - Max-file limit for pilot runs.
    - Resume mode (skip files already present in CSV).
    - Row-by-row CSV writes (safe to stop and continue later).
    - Optional lightweight mode (skip deep group recursion, no text samples).

  Usage:
    1) Open Adobe Illustrator.
    2) File > Scripts > Other Script...
    3) Run this JSX and paste a root folder path.
*/

(function () {
    var header = [
        "filename",
        "path",
        "status",
        "notes",
        "layer_count",
        "layer_names",
        "text_frame_count",
        "group_item_count",
        "path_item_count",
        "compound_path_item_count",
        "raster_item_count",
        "placed_item_count",
        "plugin_item_count",
        "mesh_item_count",
        "symbol_item_count",
        "graph_item_count",
        "legacy_text_item_count",
        "other_item_count",
        "total_page_item_count",
        "text_samples"
    ];

    function csvEscape(value) {
        if (value === undefined || value === null) {
            value = "";
        }
        value = String(value).replace(/\r/g, "\\r").replace(/\n/g, "\\n");
        if (/[",\r\n]/.test(value)) {
            value = '"' + value.replace(/"/g, '""') + '"';
        }
        return value;
    }

    function safeText(value) {
        if (value === undefined || value === null) {
            return "";
        }
        return String(value).replace(/\r/g, "\\r").replace(/\n/g, "\\n").replace(/\t/g, " ");
    }

    function rowToCsv(row) {
        var out = [];
        for (var i = 0; i < row.length; i += 1) {
            out.push(csvEscape(row[i]));
        }
        return out.join(",");
    }

    function writeCsvLine(fileObj, line, mode) {
        fileObj.encoding = "UTF-8";
        if (!fileObj.open(mode)) {
            return false;
        }
        fileObj.writeln(line);
        fileObj.close();
        return true;
    }

    function collectEps(folder, acc) {
        var entries = folder.getFiles();
        for (var i = 0; i < entries.length; i += 1) {
            var entry = entries[i];
            if (entry instanceof Folder) {
                collectEps(entry, acc);
            } else if (entry instanceof File && /\.eps$/i.test(entry.name)) {
                acc.push(entry);
            }
        }
    }

    function makeCounts() {
        return {
            layer_count: 0,
            text_frame_count: 0,
            group_item_count: 0,
            path_item_count: 0,
            compound_path_item_count: 0,
            raster_item_count: 0,
            placed_item_count: 0,
            plugin_item_count: 0,
            mesh_item_count: 0,
            symbol_item_count: 0,
            graph_item_count: 0,
            legacy_text_item_count: 0,
            other_item_count: 0,
            total_page_item_count: 0,
            layer_names: [],
            text_samples: []
        };
    }

    function bumpCount(counts, item, captureSamples) {
        counts.total_page_item_count += 1;
        var t = item.typename;
        if (t === "TextFrame") {
            counts.text_frame_count += 1;
            if (captureSamples) {
                try {
                    if (counts.text_samples.length < 10 && item.contents) {
                        counts.text_samples.push(safeText(item.contents));
                    }
                } catch (e) {
                }
            }
            return;
        }
        if (t === "GroupItem") {
            counts.group_item_count += 1;
            return;
        }
        if (t === "PathItem") {
            counts.path_item_count += 1;
            return;
        }
        if (t === "CompoundPathItem") {
            counts.compound_path_item_count += 1;
            return;
        }
        if (t === "RasterItem") {
            counts.raster_item_count += 1;
            return;
        }
        if (t === "PlacedItem") {
            counts.placed_item_count += 1;
            return;
        }
        if (t === "PluginItem") {
            counts.plugin_item_count += 1;
            return;
        }
        if (t === "MeshItem") {
            counts.mesh_item_count += 1;
            return;
        }
        if (t === "SymbolItem") {
            counts.symbol_item_count += 1;
            return;
        }
        if (t === "GraphItem") {
            counts.graph_item_count += 1;
            return;
        }
        if (t === "LegacyTextItem") {
            counts.legacy_text_item_count += 1;
            return;
        }
        counts.other_item_count += 1;
    }

    function walkItems(items, counts, deepWalk, captureSamples) {
        for (var i = 0; i < items.length; i += 1) {
            var item = items[i];
            bumpCount(counts, item, captureSamples);
            if (deepWalk && item.typename === "GroupItem") {
                walkItems(item.pageItems, counts, deepWalk, captureSamples);
            }
        }
    }

    function buildRow(fileObj, counts, status, notes) {
        return [
            fileObj.name,
            fileObj.fsName,
            status,
            notes,
            counts.layer_count,
            counts.layer_names.join(" | "),
            counts.text_frame_count,
            counts.group_item_count,
            counts.path_item_count,
            counts.compound_path_item_count,
            counts.raster_item_count,
            counts.placed_item_count,
            counts.plugin_item_count,
            counts.mesh_item_count,
            counts.symbol_item_count,
            counts.graph_item_count,
            counts.legacy_text_item_count,
            counts.other_item_count,
            counts.total_page_item_count,
            counts.text_samples.join(" | ")
        ];
    }

    function readProcessedPaths(outFile) {
        var done = {};
        if (!outFile.exists) {
            return done;
        }
        outFile.encoding = "UTF-8";
        if (!outFile.open("r")) {
            return done;
        }
        var lineNo = 0;
        while (!outFile.eof) {
            var line = outFile.readln();
            lineNo += 1;
            if (!line || lineNo === 1) {
                continue;
            }
            var m = line.match(/^[^,]*,(\"(?:[^\"]|\"\")*\"|[^,]*)/);
            if (!m) {
                continue;
            }
            var raw = m[1];
            if (raw.length >= 2 && raw.charAt(0) === '"' && raw.charAt(raw.length - 1) === '"') {
                raw = raw.substring(1, raw.length - 1).replace(/""/g, '"');
            }
            done[raw] = true;
        }
        outFile.close();
        return done;
    }

    var rootPath = prompt("Paste folder path to scan recursively for EPS:", "");
    if (!rootPath) {
        alert("No folder path provided.");
        return;
    }
    rootPath = rootPath.replace(/^"|"$/g, "");
    var rootFolder = new Folder(rootPath);
    if (!rootFolder.exists) {
        alert("Folder does not exist:\n" + rootPath);
        return;
    }

    var maxFilesRaw = prompt("Max EPS files to process (0 = all):", "0");
    var maxFiles = parseInt(maxFilesRaw, 10);
    if (isNaN(maxFiles) || maxFiles < 0) {
        maxFiles = 0;
    }

    var deepWalkAnswer = prompt("Deep group walk? yes/no (yes = slower, more complete)", "no");
    var deepWalk = /^y(es)?$/i.test(deepWalkAnswer || "");

    var sampleAnswer = prompt("Capture text samples? yes/no (yes = slower)", "no");
    var captureSamples = /^y(es)?$/i.test(sampleAnswer || "");

    var resumeAnswer = prompt("Resume from existing CSV if present? yes/no", "yes");
    var resume = !resumeAnswer || /^y(es)?$/i.test(resumeAnswer);

    var outFile = new File(rootFolder.fsName + "/illustrator_eps_probe_summary.csv");
    var alreadyDone = resume ? readProcessedPaths(outFile) : {};

    var epsFiles = [];
    collectEps(rootFolder, epsFiles);
    epsFiles.sort(function (a, b) {
        var aa = a.fsName.toLowerCase();
        var bb = b.fsName.toLowerCase();
        if (aa < bb) {
            return -1;
        }
        if (aa > bb) {
            return 1;
        }
        return 0;
    });

    if (epsFiles.length === 0) {
        alert("No EPS files found under:\n" + rootFolder.fsName);
        return;
    }

    var filtered = [];
    for (var i = 0; i < epsFiles.length; i += 1) {
        var p = epsFiles[i].fsName;
        if (!alreadyDone[p]) {
            filtered.push(epsFiles[i]);
        }
    }

    if (maxFiles > 0 && filtered.length > maxFiles) {
        filtered = filtered.slice(0, maxFiles);
    }

    if (filtered.length === 0) {
        alert("Nothing to process. All files appear to be already in CSV.");
        return;
    }

    var append = outFile.exists && resume;
    if (!append) {
        if (!writeCsvLine(outFile, rowToCsv(header), "w")) {
            alert("Cannot open output CSV:\n" + outFile.fsName);
            return;
        }
    }

    var errors = 0;
    for (var fileIndex = 0; fileIndex < filtered.length; fileIndex += 1) {
        var fileObj = filtered[fileIndex];
        var status = "OK";
        var notes = "";
        var doc = null;
        var counts = makeCounts();
        try {
            $.writeln("Probing " + (fileIndex + 1) + "/" + filtered.length + ": " + fileObj.fsName);
            doc = app.open(fileObj);
            counts.layer_count = doc.layers.length;

            for (var layerIndex = 0; layerIndex < doc.layers.length; layerIndex += 1) {
                var layer = doc.layers[layerIndex];
                counts.layer_names.push(safeText(layer.name));
                walkItems(layer.pageItems, counts, deepWalk, captureSamples);
            }
        } catch (e) {
            status = "ERROR";
            notes = safeText(e);
            errors += 1;
        } finally {
            if (doc !== null) {
                try {
                    doc.close(SaveOptions.DONOTSAVECHANGES);
                } catch (closeErr) {
                    if (status === "OK") {
                        status = "CLOSE ERROR";
                    }
                    notes = (notes ? notes + " | " : "") + "close: " + safeText(closeErr);
                }
            }
        }

        if (!writeCsvLine(outFile, rowToCsv(buildRow(fileObj, counts, status, notes)), "a")) {
            alert("Cannot append to output CSV:\n" + outFile.fsName);
            return;
        }
    }
    alert(
        "Done.\n" +
        "Processed: " + filtered.length + "\n" +
        "Errors: " + errors + "\n" +
        "CSV: " + outFile.fsName
    );
}());
