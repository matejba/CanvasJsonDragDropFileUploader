# CanvasJsonDragDropFileUploader (PCF)

A PCF control for Canvas Apps that lets users add files from their device into the app using:
- drag & drop into the dropzone
- clicking the dropzone and selecting files via the file picker dialog

The control is fully responsive and shows a list of selected files. Each file can be:
- removed (X icon)
- renamed (pencil icon)

Note: The control only loads files into the Canvas App as JSON output. Actual file upload/storage must be implemented separately (for example using Power Automate).

---

## Parameters

### DropText
Custom text displayed inside the dropzone.

- Default: `Drag and drop files, or click to upload`

---

### MaxFiles
Maximum number of files that can be added to this control instance.

- Default: `0` (no uploads allowed until increased)
- Example: `10` allows up to 10 files

---

### AllowedFormats
Defines which file formats are allowed.

Provide one or more file extensions separated by commas (each extension must start with a dot).

Examples:
- `.xlsx,.docx`
- `.xlsx`

If left empty, any file format is allowed.

---

### Reset
Clears all currently selected files.

To trigger reset, change the value to anything different than the previous value (recommended: bind it to a variable and update it when you want to clear the control).

---

### Save button color
Sets the rename dialog **Save** button color in HEX format.

Examples:
- `#000000` (black)

If left empty, the default blue color is used.

---

## Output

### filesJson
JSON array containing all selected files. Each item includes:

- `name` (file name)
- `size` (file size in bytes)
- `type` (MIME content type)
- `lastModified` (timestamp)
- `content` (Base64 file content)

Because the output is valid JSON, it can be sent directly to Power Automate for further processing (creating files, uploading to SharePoint, etc.).

---

## User actions

- Click **X** to remove a selected file from the list
- Click the **pencil icon** to rename a file (extension is preserved automatically)# CanvasJsonDragDropFileUploader
