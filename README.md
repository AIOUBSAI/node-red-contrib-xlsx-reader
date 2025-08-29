# node-red-contrib-xlsx-reader

Node-RED node to read Excel files/directories, aggregate sheets into a single output, and optionally fill “merged-like” columns.

## Features

* Read a **single file** or a **directory** of `.xlsx` files.
* Skip sheets by **regex** and/or **hidden** flag.
* **One aggregated output message** (no per-sheet spam).
* Choose where to store the result: **`msg`**, **`flow`**, or **`global`** + a **deep path** (e.g. `data`, `store.xl`).
* Optional **fill-forward** for specified columns to handle Excel **merged cells** exported as blanks.

## Install

From your Node-RED user dir (`~/.node-red`):

```bash
npm install node-red-contrib-xlsx-reader
```

Or via **Manage palette** → **Install** → search for `node-red-contrib-xlsx-reader`.

## Usage

1. Drag **“xlsx-reader”** into your flow.
2. Set **Path** to a file or directory.

   * Use typed input to choose `str`, `msg`, `flow`, `global`, or `env`.
3. Choose **Mode**: `Single File` or `Directory`.
4. Optionally set:

   * **Exclude Regex** (e.g. `^\.` to skip sheets starting with a dot).
   * **Include Hidden** (to include hidden/veryHidden sheets).
5. Configure **Output Target**:

   * **Scope**: `msg`, `flow`, or `global`.
   * **Path**: deep path where data is stored (e.g. `data`, `store.xl`).
6. (Optional) **Fill Merged-like Cells**:

   * Check the box and list **column headers** (comma-separated) to forward-fill blank cells.

### Output Structure

If `Output Target` is `msg` and `Path` is `data`, the node writes:

```js
msg.data = {
  data: {
    "/abs/path/file1.xlsx": {
      "Sheet1": [ { ...row }, ... ],
      "Sheet2": [ { ...row }, ... ]
    },
    "/abs/path/file2.xlsx": {
      "SheetA": [ { ...row }, ... ]
    }
  },
  summary: { fileCount: 2, sheetCount: 3, rowCount: 1234 },
  options: {
    excludeRegex: "^\\.",
    includeHidden: false,
    fillMerged: true,
    mergedColumns: ["Family","CategoryID"]
  }
};
return msg;
```

> Tip: If you see columns like `__EMPTY`, it usually means there are **more columns than headers** in the sheet. You can clean them downstream or align headers in the source file.

## Properties

| Property                   | Type                            | Description                                                                           |
| -------------------------- | ------------------------------- | ------------------------------------------------------------------------------------- |
| **Path**                   | str / msg / flow / global / env | File or directory path to read.                                                       |
| **Mode**                   | enum                            | `file` or `directory`.                                                                |
| **Exclude Regex**          | string                          | Skip sheets whose names match this regex.                                             |
| **Include Hidden**         | boolean                         | Include hidden/veryHidden sheets if checked.                                          |
| **Output Target**          | scope + path                    | Where to store the aggregated result (choose `msg`, `flow`, or `global` + deep path). |
| **Fill Merged-like Cells** | boolean                         | If checked, forward-fill blanks for specified columns.                                |
| **Columns to Fill**        | string                          | Comma-separated **header names** (case-sensitive).                                    |

## Example Flow

```json
[
  {
    "id": "inject1",
    "type": "inject",
    "z": "flow1",
    "name": "Tick",
    "props": [],
    "repeat": "",
    "once": true
  },
  {
    "id": "xlsx1",
    "type": "xlsx-reader",
    "z": "flow1",
    "name": "Read XLSX dir",
    "path": "/path/to/dir/of/xlsx",
    "pathType": "str",
    "mode": "directory",
    "excludeRegex": "^\\.",
    "includeHidden": false,
    "outputTargetType": "msg",
    "outputTargetPath": "data",
    "fillMerged": true,
    "mergedColumns": "EEID,Family"
  },
  {
    "id": "debug1",
    "type": "debug",
    "z": "flow1",
    "name": "See msg.data",
    "active": true,
    "tosidebar": true,
    "complete": "data",
    "statusVal": ""
  },
  {
    "id": "link1",
    "type": "link",
    "source": "inject1",
    "target": "xlsx1"
  },
  {
    "id": "link2",
    "type": "link",
    "source": "xlsx1",
    "target": "debug1"
  }
]
```

*(You’ll still set node UI options in the editor; this is just a sketch.)*

## Notes

* **Performance**: this node aggregates all sheets into memory for one output. For very large workbooks, consider processing in batches.
* **Headers**: forward-fill uses **header names** exactly as parsed by `xlsx`. Ensure consistent casing.
* **Icons**: the node uses the default `file.png`. To use a custom icon, add `icons/youricon.svg` and reference it in the HTML.

## Changelog

* **0.1.0** – Initial public release: directory/single-file modes, hidden/regex filters, aggregated output, msg/flow/global target, merged-cells fill-forward.

## License

MIT © AIOUBSAI

---

### Next steps (repo)

1. Put `package.json`, `xlsx-reader.js`, `xlsx-reader.html`, `README.md`, and `LICENSE` **in the repo root** (as you asked).
2. Commit & push to GitHub.
3. Ensure `package.json` has:

   * `"license": "MIT"`
   * `"keywords": ["node-red", ...]`
   * `"repository"`, `"bugs"`, `"homepage"` pointing to your GitHub repo.

When you’re ready, I can check your `package.json` and tweak it for publish-readiness.
