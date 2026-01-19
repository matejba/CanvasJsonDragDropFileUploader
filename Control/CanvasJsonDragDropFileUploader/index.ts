import { IInputs, IOutputs } from "./generated/ManifestTypes";

interface FileDescriptor {
    name: string;
    size: number;
    type: string;
    lastModified: number;
    content: string;
}

export class CanvasJsonDragDropFileUploader implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private container: HTMLDivElement;
    private fileInput: HTMLInputElement;
    private files: FileDescriptor[] = [];
    private notifyOutputChanged: () => void;
    private maxFiles: number | null = null;
    private allowedFormats: string | null = null;
    private saveButtonColor: string | null = null;
    private labelElement: HTMLDivElement;
    private filesListElement: HTMLDivElement;
    private progressElement: HTMLDivElement;
    private errorElement: HTMLDivElement;
    private isLoading = false;
    private errorTimeoutId: number | null = null;

    private lastResetToken: number | null = null;

    private renameDialogBackdrop: HTMLDivElement | null = null;
    private renameDialog: HTMLDivElement | null = null;
    private renameInput: HTMLInputElement | null = null;
    private renameExtLabel: HTMLSpanElement | null = null;
    private renameError: HTMLDivElement | null = null;
    private renameOkButton: HTMLButtonElement | null = null;
    private currentRenameIndex: number | null = null;
    private currentRenameExt = "";
    private clearFiles(): void {
        console.log("clearFiles");
        this.files = [];

        // reset input so the same file can be selected again
        if (this.fileInput) {
            this.fileInput.value = "";
        }

        this.setError(null);
        this.renderFilesList();
        this.notifyOutputChanged();
    }

    private getFileExt(filename: string): string {
        const name = (filename || "").trim();
        const lastDot = name.lastIndexOf(".");
        if (lastDot <= 0) return "";
        return name.substring(lastDot); // includes dot
    }

    private getFileBase(filename: string): string {
        const name = (filename || "").trim();
        const lastDot = name.lastIndexOf(".");
        if (lastDot <= 0) return name;
        return name.substring(0, lastDot);
    }

    private applyRenameOkButtonStyle(): void {
    if (!this.renameOkButton) return;

    const color = (this.saveButtonColor || "").trim();
    const finalColor = color.length > 0 ? color : "#0078d4";

    this.renameOkButton.style.border = `1px solid ${finalColor}`;
    this.renameOkButton.style.background = finalColor;
}

    private ensureRenameDialog(): void {
        if (this.renameDialogBackdrop) return;

        const backdrop = document.createElement("div");
        backdrop.style.position = "fixed";
        backdrop.style.inset = "0";
        backdrop.style.background = "rgba(0,0,0,0.35)";
        backdrop.style.display = "none";
        backdrop.style.alignItems = "center";
        backdrop.style.justifyContent = "center";
        backdrop.style.zIndex = "9999";

        const dialog = document.createElement("div");
        dialog.style.width = "min(440px, calc(100vw - 32px))";
        dialog.style.background = "#fff";
        dialog.style.borderRadius = "10px";
        dialog.style.boxShadow = "0 10px 30px rgba(0,0,0,0.2)";
        dialog.style.padding = "14px";
        dialog.style.boxSizing = "border-box";
        dialog.style.fontFamily = "Segoe UI, system-ui, -apple-system, BlinkMacSystemFont, 'Helvetica Neue', Arial, sans-serif";
        dialog.addEventListener("click", e => e.stopPropagation());

        const title = document.createElement("div");
        title.innerText = "Rename file";
        title.style.fontSize = "13px";
        title.style.fontWeight = "600";
        title.style.marginBottom = "10px";
        title.style.fontFamily = "inherit";

        const row = document.createElement("div");
        row.style.display = "flex";
        row.style.alignItems = "center";
        row.style.columnGap = "8px";

        const input = document.createElement("input");
        input.type = "text";
        input.style.flex = "1 1 auto";
        input.style.height = "34px";
        input.style.border = "1px solid #c8c8c8";
        input.style.borderRadius = "8px";
        input.style.padding = "0 10px";
        input.style.fontSize = "13px";
        input.style.outline = "none";
        input.style.fontFamily = "inherit";

        const ext = document.createElement("span");
        ext.style.flex = "0 0 auto";
        ext.style.fontSize = "12px";
        ext.style.color = "#666";
        ext.style.padding = "0 10px";
        ext.style.height = "34px";
        ext.style.display = "inline-flex";
        ext.style.alignItems = "center";
        ext.style.border = "1px solid #e0e0e0";
        ext.style.borderRadius = "8px";
        ext.style.background = "#fafafa";
        ext.style.fontFamily = "inherit";

        row.appendChild(input);
        row.appendChild(ext);

        const err = document.createElement("div");
        err.style.marginTop = "8px";
        err.style.fontSize = "12px";
        err.style.color = "#c00000";
        err.style.display = "none";

        const actions = document.createElement("div");
        actions.style.display = "flex";
        actions.style.justifyContent = "flex-end";
        actions.style.columnGap = "8px";
        actions.style.marginTop = "12px";

        const btnCancel = document.createElement("button");
        btnCancel.type = "button";
        btnCancel.innerText = "Cancel";
        btnCancel.style.height = "34px";
        btnCancel.style.padding = "0 12px";
        btnCancel.style.borderRadius = "8px";
        btnCancel.style.border = "1px solid #d0d0d0";
        btnCancel.style.background = "#fff";
        btnCancel.style.cursor = "pointer";
        btnCancel.style.fontFamily = "inherit";

        const btnOk = document.createElement("button");
        btnOk.type = "button";
        btnOk.innerText = "Save";
        btnOk.style.height = "34px";
        btnOk.style.padding = "0 12px";
        btnOk.style.borderRadius = "8px";
        btnOk.style.border = "1px solid #0078d4";
        btnOk.style.background = "#0078d4";
        btnOk.style.color = "#fff";
        btnOk.style.cursor = "pointer";
        btnOk.style.fontFamily = "inherit";
        this.renameOkButton = btnOk;
this.applyRenameOkButtonStyle();

        actions.appendChild(btnCancel);
        actions.appendChild(btnOk);

        dialog.appendChild(title);
        dialog.appendChild(row);
        dialog.appendChild(err);
        dialog.appendChild(actions);

        backdrop.appendChild(dialog);
        document.body.appendChild(backdrop);

        const close = () => this.closeRenameDialog();
        const submit = () => this.submitRenameDialog();

        backdrop.addEventListener("click", close);
        btnCancel.addEventListener("click", close);
        btnOk.addEventListener("click", submit);

        input.addEventListener("keydown", e => {
            if (e.key === "Enter") {
                e.preventDefault();
                submit();
            }
            if (e.key === "Escape") {
                e.preventDefault();
                close();
            }
        });

        this.renameDialogBackdrop = backdrop;
        this.renameDialog = dialog;
        this.renameInput = input;
        this.renameExtLabel = ext;
        this.renameError = err;
    }

    private openRenameDialog(index: number): void {
        const current = this.files[index];
        if (!current) return;

        this.ensureRenameDialog();

        this.currentRenameIndex = index;
        this.currentRenameExt = this.getFileExt(current.name);

        const base = this.getFileBase(current.name);

        if (this.renameInput) {
            this.renameInput.value = base;
        }
        if (this.renameExtLabel) {
            this.renameExtLabel.innerText = this.currentRenameExt || "(no extension)";
            this.renameExtLabel.style.display = this.currentRenameExt ? "inline-flex" : "inline-flex";
        }
        if (this.renameError) {
            this.renameError.innerText = "";
            this.renameError.style.display = "none";
        }

        if (this.renameDialogBackdrop) {
            this.renameDialogBackdrop.style.display = "flex";
        }

        window.setTimeout(() => {
            this.renameInput?.focus();
            this.renameInput?.select();
        }, 0);
    }

    private closeRenameDialog(): void {
        if (this.renameDialogBackdrop) {
            this.renameDialogBackdrop.style.display = "none";
        }
        this.currentRenameIndex = null;
        this.currentRenameExt = "";
        if (this.renameError) {
            this.renameError.innerText = "";
            this.renameError.style.display = "none";
        }
    }

    private submitRenameDialog(): void {
        if (this.currentRenameIndex === null) return;

        const current = this.files[this.currentRenameIndex];
        if (!current) {
            this.closeRenameDialog();
            return;
        }

        const rawBase = (this.renameInput?.value || "").trim();

        let base = rawBase;
        if (this.currentRenameExt) {
            const dot = base.lastIndexOf(".");
            if (dot > 0) {
                base = base.substring(0, dot);
            }
            base = base.replace(/[.\s]+$/g, "").trim();
        }

        if (!base) {
            if (this.renameError) {
                this.renameError.innerText = "Enter a file name.";
                this.renameError.style.display = "block";
            }
            return;
        }

        const nextName = this.currentRenameExt ? `${base}${this.currentRenameExt}` : base;

        console.log("rename file", {
            index: this.currentRenameIndex,
            from: current.name,
            to: nextName
        });

        current.name = nextName;
        this.renderFilesList();
        this.notifyOutputChanged();
        this.closeRenameDialog();
    }

    private createIconButton(title: string, svgPathD: string): HTMLButtonElement {
        const btn = document.createElement("button");
        btn.type = "button";
        btn.title = title;
        btn.setAttribute("aria-label", title);

        btn.style.width = "30px";
        btn.style.height = "30px";
        btn.style.display = "inline-flex";
        btn.style.alignItems = "center";
        btn.style.justifyContent = "center";
        btn.style.border = "1px solid transparent";
        btn.style.borderRadius = "8px";
        btn.style.background = "transparent";
        btn.style.cursor = "pointer";
        btn.style.padding = "0";

        btn.addEventListener("mouseenter", () => {
            btn.style.background = "#f4f4f4";
            btn.style.borderColor = "#e5e5e5";
        });
        btn.addEventListener("mouseleave", () => {
            btn.style.background = "transparent";
            btn.style.borderColor = "transparent";
        });

        const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
        svg.setAttribute("viewBox", "0 0 24 24");
        svg.setAttribute("width", "18");
        svg.setAttribute("height", "18");

        const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
        path.setAttribute("d", svgPathD);
        path.setAttribute("fill", "none");
        path.setAttribute("stroke", "#555");
        path.setAttribute("stroke-width", "2.2");
        path.setAttribute("stroke-linecap", "round");
        path.setAttribute("stroke-linejoin", "round");

        svg.appendChild(path);
        btn.appendChild(svg);

        return btn;
    }

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        console.log("CanvasFileDropzone.init start");

        this.notifyOutputChanged = notifyOutputChanged;

        this.container = document.createElement("div");
        this.container.style.display = "flex";
        this.container.style.flexDirection = "column";
        this.container.style.alignItems = "stretch";
        this.container.style.justifyContent = "flex-start";
        this.container.style.boxSizing = "border-box";
        this.container.style.height = "100%";
        this.container.style.minHeight = "0";

        const dropZone = document.createElement("div");
        dropZone.style.border = "2px dashed #888";
        dropZone.style.borderRadius = "8px";
        dropZone.style.display = "flex";
        dropZone.style.alignItems = "center";
        dropZone.style.justifyContent = "center";
        dropZone.style.padding = "16px";
        dropZone.style.minHeight = "80px";
        dropZone.style.boxSizing = "border-box";
        dropZone.style.cursor = "pointer";

        this.labelElement = document.createElement("div");
        this.labelElement.innerText =
            context.parameters.dropText.raw || "Drag and drop files, or click to upload";
        this.labelElement.style.textAlign = "center";

        dropZone.appendChild(this.labelElement);

        this.progressElement = document.createElement("div");
        this.progressElement.style.marginTop = "6px";
        this.progressElement.style.height = "4px";
        this.progressElement.style.width = "100%";
        this.progressElement.style.background = "#eee";
        this.progressElement.style.borderRadius = "2px";
        this.progressElement.style.overflow = "hidden";
        this.progressElement.style.display = "none";

        const progressInner = document.createElement("div");
        progressInner.style.width = "100%";
        progressInner.style.height = "100%";
        progressInner.style.background = "#0078d4";
        progressInner.style.opacity = "0.6";
        this.progressElement.appendChild(progressInner);

        this.errorElement = document.createElement("div");
        this.errorElement.style.marginTop = "4px";
        this.errorElement.style.fontSize = "11px";
        this.errorElement.style.color = "#c00000";
        this.errorElement.style.display = "none";

        this.filesListElement = document.createElement("div");
        this.filesListElement.style.marginTop = "8px";
        this.filesListElement.style.display = "flex";
        this.filesListElement.style.flexDirection = "column";
        this.filesListElement.style.rowGap = "4px";
        this.filesListElement.style.flex = "1 1 auto";
        this.filesListElement.style.minHeight = "0";
        this.filesListElement.style.overflowY = "auto";

        this.fileInput = document.createElement("input");
        this.fileInput.type = "file";
        this.fileInput.multiple = true;
        this.fileInput.style.display = "none";

        this.container.appendChild(dropZone);
        this.container.appendChild(this.progressElement);
        this.container.appendChild(this.errorElement);
        this.container.appendChild(this.filesListElement);
        container.appendChild(this.container);
        container.appendChild(this.fileInput);

        this.ensureRenameDialog();
        this.renderFilesList();

        dropZone.addEventListener("click", () => {
            console.log("dropZone click");
            if (!this.fileInput) {
                console.warn("fileInput not defined on click");
                return;
            }
            this.fileInput.click();
        });

        dropZone.addEventListener("dragover", e => {
            e.preventDefault();
            console.log("dropZone dragover");
        });

        dropZone.addEventListener("drop", e => {
            e.preventDefault();
            console.log("dropZone drop", { hasFiles: !!e.dataTransfer?.files });
            if (e.dataTransfer?.files) {
                this.handleFiles(e.dataTransfer.files);
            }
        });

        this.fileInput.addEventListener("change", () => {
            console.log("fileInput change", {
                filesCount: this.fileInput.files ? this.fileInput.files.length : 0
            });
            if (this.fileInput.files) {
                this.handleFiles(this.fileInput.files);
            }
        });

        console.log("CanvasFileDropzone.init end");
    }

    private setLoading(loading: boolean): void {
        this.isLoading = loading;
        if (this.progressElement) {
            this.progressElement.style.display = loading ? "block" : "none";
        }
        console.log("setLoading", { loading });
    }

    private setError(message: string | null): void {
        if (!this.errorElement) {
            return;
        }

        if (this.errorTimeoutId !== null) {
            window.clearTimeout(this.errorTimeoutId);
            this.errorTimeoutId = null;
        }

        if (message && message.length > 0) {
            this.errorElement.innerText = message;
            this.errorElement.style.display = "block";

            this.errorTimeoutId = window.setTimeout(() => {
                this.errorElement.innerText = "";
                this.errorElement.style.display = "none";
                this.errorTimeoutId = null;
            }, 5000);
        } else {
            this.errorElement.innerText = "";
            this.errorElement.style.display = "none";
        }

        console.log("setError", { message });
    }

    private isFileAllowed(file: File): boolean {
        if (!this.allowedFormats || this.allowedFormats.trim().length === 0) {
            return true;
        }

        const spec = this.allowedFormats
            .split(",")
            .map(s => s.trim().toLowerCase())
            .filter(s => s.length > 0);

        if (spec.length === 0) {
            return true;
        }

        const mime = (file.type || "").toLowerCase();
        const name = file.name.toLowerCase();

        for (const token of spec) {
            if (token === "*") return true;

            if (token.endsWith("/*")) {
                const prefix = token.replace("/*", "");
                if (mime.startsWith(prefix + "/")) return true;
            } else if (token.startsWith(".")) {
                if (name.endsWith(token)) return true;
            } else if (token.includes("/")) {
                if (mime === token) return true;
            }
        }

        return false;
    }

    private handleFiles(fileList: FileList) {
        console.log("handleFiles start", { fileCount: fileList.length });
        this.setError(null);

        let incoming = Array.from(fileList);

        const disallowed = incoming.filter(f => !this.isFileAllowed(f));
        if (disallowed.length > 0) {
            const disallowedCount = disallowed.length;
            this.setError(`${disallowedCount === 1 ? "A file was not uploaded" : "Some files were not uploaded"}. Allowed formats: ${this.allowedFormats || "any"}.`);
            incoming = incoming.filter(f => this.isFileAllowed(f));
        }

        let newFiles = incoming;

        if (this.maxFiles !== null) {
            const remaining = this.maxFiles - this.files.length;
            console.log("handleFiles maxFiles check", {
                maxFiles: this.maxFiles,
                existing: this.files.length,
                remaining
            });

            if (remaining <= 0) {
                this.setError(`No more files can be added. Maximum number of files is ${this.maxFiles}.`);                
                console.warn("handleFiles: no remaining slots for files");
                return;
            }

            if (newFiles.length > remaining) {
                this.setError(`Some files were not uploaded. Maximum number of files is ${this.maxFiles}.`);
                newFiles = newFiles.slice(0, remaining);
            }
        }

        if (newFiles.length === 0) {
            console.warn("handleFiles: newFiles is empty after filtering");
            return;
        }

        const tasks = newFiles.map(f => this.readFile(f));
        console.log("handleFiles: created read tasks", { tasksCount: tasks.length });

        this.setLoading(true);

        return Promise.all(tasks)
            .then(descriptors => {
                console.log("handleFiles: Promise.all resolved", { descriptorsCount: descriptors.length });
                this.files = [...this.files, ...descriptors];
                this.renderFilesList();
                this.notifyOutputChanged();
                return null;
            })
            .catch(error => {
                console.error("handleFiles error", error);
            })
            .finally(() => {
                this.setLoading(false);
                console.log("handleFiles end");
            });
    }

    private renderFilesList(): void {
        console.log("renderFilesList", { filesCount: this.files.length });

        if (!this.filesListElement) {
            console.warn("renderFilesList: filesListElement missing");
            return;
        }

        this.filesListElement.innerHTML = "";

        const header = document.createElement("div");
        header.style.fontSize = "11px";
        header.style.color = "#555";
        const count = this.files.length;
        header.innerText = count === 0 ? "No files selected" : count === 1 ? "1 file selected" : `${count} files selected`;        
        this.filesListElement.appendChild(header);

        const pencilPath =
            "M12 20h9 M16.5 3.5a2.1 2.1 0 0 1 3 3L8 18l-4 1 1-4 11.5-11.5z";
        const xPath =
            "M18 6L6 18 M6 6l12 12";

        this.files.forEach((file, index) => {
            const row = document.createElement("div");
            row.style.display = "flex";
            row.style.alignItems = "center";
            row.style.justifyContent = "space-between";
            row.style.fontSize = "12px";
            row.style.color = "#333";
            row.style.columnGap = "8px";

            const nameSpan = document.createElement("span");
            nameSpan.innerText = file.name;
            nameSpan.style.overflow = "hidden";
            nameSpan.style.textOverflow = "ellipsis";
            nameSpan.style.whiteSpace = "nowrap";
            nameSpan.style.flex = "1 1 auto";

            const actions = document.createElement("div");
            actions.style.display = "flex";
            actions.style.alignItems = "center";
            actions.style.columnGap = "6px";
            actions.style.flex = "0 0 auto";

            const renameButton = this.createIconButton("Rename", pencilPath);
            renameButton.addEventListener("click", () => {
                console.log("rename file click", { index, name: file.name });
                this.openRenameDialog(index);
            });

            const removeButton = this.createIconButton("Remove", xPath);
            removeButton.addEventListener("click", () => {
                console.log("remove file click", { index, name: file.name });
                this.files.splice(index, 1);
                this.renderFilesList();
                this.notifyOutputChanged();
            });

            actions.appendChild(renameButton);
            actions.appendChild(removeButton);

            row.appendChild(nameSpan);
            row.appendChild(actions);

            this.filesListElement.appendChild(row);
        });
    }

    private readFile(file: File): Promise<FileDescriptor> {
        console.log("readFile", { name: file.name, size: file.size, type: file.type });

        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => {
                const result = reader.result as string;
                const base64 = result.split(",")[1] || "";
                resolve({
                    name: file.name,
                    size: file.size,
                    type: file.type,
                    lastModified: file.lastModified,
                    content: base64
                });
            };
            reader.onerror = () => {
                console.error("readFile error", reader.error);
                reject(reader.error);
            };
            reader.readAsDataURL(file);
        });
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this.maxFiles = context.parameters.maxFiles.raw != null ? Number(context.parameters.maxFiles.raw) : 10;
        this.allowedFormats = context.parameters.allowedFormats.raw || null;
        this.saveButtonColor = context.parameters.saveButtonColor?.raw || null;
this.applyRenameOkButtonStyle();

        const resetToken = context.parameters.reset?.raw ?? null;
        if (resetToken !== null && resetToken !== this.lastResetToken) {
            console.log("reset token changed", { from: this.lastResetToken, to: resetToken });
            this.lastResetToken = resetToken;
            if (this.files.length > 0) {
                this.clearFiles();
            }
        }

        const dropText = context.parameters.dropText.raw;
        if (this.labelElement) {
            this.labelElement.innerText =
                dropText && dropText.length > 0 ? dropText : "Drag and drop files, or click to upload";
        }

        if (this.fileInput) {
            this.fileInput.accept = this.allowedFormats && this.allowedFormats.length > 0 ? this.allowedFormats : "";
        }

        if (this.container) {
            const width = context.mode.allocatedWidth;
            const height = context.mode.allocatedHeight;
            if (width > 0) this.container.style.width = `${width}px`;
            if (height > 0) this.container.style.height = `${height}px`;
        }

        console.log("updateView", {
            maxFiles: this.maxFiles,
            allowedFormats: this.allowedFormats,
            width: context.mode.allocatedWidth,
            height: context.mode.allocatedHeight
        });
    }

    public getOutputs(): IOutputs {
        const json = JSON.stringify(this.files);
        console.log("getOutputs", { length: this.files.length });
        return { filesJson: json };
    }

    public destroy(): void {
        console.log("CanvasFileDropzone.destroy");
        if (this.renameDialogBackdrop && this.renameDialogBackdrop.parentElement) {
            this.renameDialogBackdrop.parentElement.removeChild(this.renameDialogBackdrop);
        }
        this.renameDialogBackdrop = null;
        this.renameDialog = null;
        this.renameInput = null;
        this.renameExtLabel = null;
        this.renameError = null;
        this.currentRenameIndex = null;
        this.currentRenameExt = "";
    }
}