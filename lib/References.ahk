; ============================================================
; lib/References.ahk -- Reference (URL / file) manager
; ------------------------------------------------------------
; References live in Prefs.references[]. URLs are restricted to
; http/https. Files are copied to a References\ subfolder using
; a sanitized destination name (no path traversal, no control
; characters). Allowed extensions: PDF, DOC(X), XLS(X), PPT(X).
; Each open increments a "uses" counter; the submenu sorts by
; that count.
; ============================================================

#Requires AutoHotkey v2.0
#Include Modern.ahk
#Include Util.ahk

global g_MaxReferencesInMenu := 15

; ---- right-click submenu ---------------------------------------------------

BuildReferencesMenu() {
    m := Menu()
    m.Add("Add Reference...",  (*) => ShowAddReferenceDialog())
    m.Add("Manage References...", (*) => ShowReferencesManager())
    m.Add()

    refs := Prefs.References()
    if (refs.Length = 0)
        return m   ; just Add/Manage on top -- no awkward placeholder

    sorted := SortReferencesByUses(refs)
    count := 0
    for ref in sorted {
        if (count >= g_MaxReferencesInMenu)
            break
        if !(ref is Map) || !ref.Has("name")
            continue
        name := ref["name"]
        m.Add(name, MakeOpenReferenceCallback(name))
        count++
    }
    return m
}

SortReferencesByUses(refs) {
    sorted := []
    for ref in refs {
        inserted := false
        uses := ref.Has("uses") ? Integer(ref["uses"]) : 0
        for i, existing in sorted {
            existingUses := existing.Has("uses") ? Integer(existing["uses"]) : 0
            if (uses > existingUses) {
                sorted.InsertAt(i, ref)
                inserted := true
                break
            }
        }
        if !inserted
            sorted.Push(ref)
    }
    return sorted
}

MakeOpenReferenceCallback(name) {
    return (*) => OpenReference(name)
}

; ---- open a reference -------------------------------------------------------

OpenReference(name) {
    idx := Prefs.FindReferenceIndex(name)
    if (idx = 0)
        return
    ref := Prefs.References()[idx]
    type := ref["type"], path := ref["path"]

    launched := false
    if (type = "url") {
        url := NormalizeURL(path)
        if (url = "") {
            MsgBox("Refusing to open invalid URL: " path, "RightClick", "Iconx")
            return
        }
        try {
            Run(url)
            launched := true
        } catch as e {
            MsgBox("Could not open URL.`n`n" url "`n`nReason: " e.Message
                , "RightClick", "Iconx")
        }
    } else {
        if !FileExist(path) {
            MsgBox("File not found:`n" path, "RightClick", "Iconx")
            return
        }
        SplitPath path, , , &ext
        if !IsValidFileType(ext) {
            MsgBox("Refusing to open file with unsupported extension: ." ext
                , "RightClick", "Iconx")
            return
        }
        try {
            Run(path)
            launched := true
        } catch as e {
            MsgBox("Could not open file.`n`n" path "`n`nReason: " e.Message
                , "RightClick", "Iconx")
        }
    }
    if launched {
        ref["uses"] := (ref.Has("uses") ? Integer(ref["uses"]) : 0) + 1
        Prefs.Save()
    }
}

; ---- add-reference dialog --------------------------------------------------

ShowAddReferenceDialog() {
    dark := Prefs.Get("display", "darkMode", false)
    palette := ModernPalette(dark)

    MouseGetPos(&mx, &my)
    work := GetWorkAreaAt(mx, my)
    w := 420, h := 220
    ; w/h are logical units (Gui.Show scales them); clamp with the
    ; physical footprint so the window stays on-screen at >100% scaling.
    scale := A_ScreenDPI / 96
    pos := ClampToWorkArea(mx + 12, my + 12, Round(w * scale), Round(h * scale), work)

    g := Gui("+AlwaysOnTop -MaximizeBox -MinimizeBox", "Add Reference")
    g.MarginX := 16, g.MarginY := 14
    g.BackColor := palette["bg"]
    g.SetFont("s10 c" palette["fg"], ModernFont())

    g.Add("Text", "x16 y14 w380", "Name:")
    eName := g.Add("Edit"
        , "x16 y32 w380 vRefName +Background" palette["bgAlt"] " c" palette["fg"])

    g.Add("Text", "x16 y66 w380", "URL (https://...) or file path:")
    ePath := g.Add("Edit"
        , "x16 y84 w300 vRefPath +Background" palette["bgAlt"] " c" palette["fg"])

    btnBrowse := g.Add("Button"
        , "x324 y82 w72 h26 +Background" palette["btnBg"] " c" palette["btnFg"]
        , "Browse...")
    btnBrowse.OnEvent("Click", (*) => BrowseReferenceFile(ePath, eName))

    btnCancel := g.Add("Button"
        , "x230 y150 w80 h32 +Background" palette["btnBg"] " c" palette["btnFg"]
        , "Cancel")
    btnCancel.OnEvent("Click", (*) => g.Destroy())

    btnSave := g.Add("Button"
        , "Default x316 y150 w80 h32 +Background" palette["btnBg"] " c" palette["btnFg"]
        , "Save")
    btnSave.OnEvent("Click", (*) => SaveReferenceFromDialog(g))

    g.OnEvent("Close",  (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())

    ApplyModernChrome(g, dark)
    g.Show("x" pos.x " y" pos.y " w" w " h" h)
    WinActivate("ahk_id " g.Hwnd)
}

BrowseReferenceFile(pathCtl, nameCtl) {
    selected := FileSelect(3, , "Select reference file"
        , "Documents (*.pdf;*.doc;*.docx;*.xls;*.xlsx;*.ppt;*.pptx)")
    if (selected = "")
        return
    pathCtl.Value := selected
    if (nameCtl.Value = "") {
        SplitPath selected, &fname
        nameCtl.Value := fname
    }
}

SaveReferenceFromDialog(g) {
    saved := g.Submit(false)
    name := Trim(saved.RefName)
    path := Trim(saved.RefPath)

    if (name = "") {
        MsgBox("Please enter a reference name.", "Add Reference", "Iconi")
        return
    }
    if (path = "") {
        MsgBox("Please enter a URL or file path.", "Add Reference", "Iconi")
        return
    }

    ; sanitize the name we use as a map key / menu label
    cleanName := RegExReplace(name, "[\x00-\x1F]", "")
    cleanName := SubStr(Trim(cleanName), 1, 100)
    if (cleanName = "") {
        MsgBox("Name contains no usable characters.", "Add Reference", "Iconx")
        return
    }

    if (Prefs.FindReferenceIndex(cleanName) != 0) {
        if (MsgBox("A reference named '" cleanName "' already exists. Replace it?"
                , "Add Reference", "YesNo Iconi") != "Yes")
            return
        Prefs.References().RemoveAt(Prefs.FindReferenceIndex(cleanName))
    }

    if FileExist(path) {
        ; --- file
        SplitPath path, &fname, , &ext
        if !IsValidFileType(ext) {
            MsgBox("Unsupported extension '." ext "'. Allowed: PDF, DOC, DOCX, XLS, XLSX, PPT, PPTX."
                , "Add Reference", "Iconx")
            return
        }
        refsDir := A_ScriptDir "\References"
        if !DirExist(refsDir)
            DirCreate refsDir
        dest := GetUniqueFilePath(refsDir, fname)
        if (dest = "") {
            MsgBox("Could not derive a safe destination filename.", "Add Reference", "Iconx")
            return
        }
        try {
            FileCopy(path, dest, 0)
        } catch as e {
            MsgBox("File copy failed: " e.Message, "Add Reference", "Iconx")
            return
        }
        Prefs.References().Push(Map(
            "name", cleanName,
            "type", "file",
            "path", dest,
            "uses", 0
        ))
    } else {
        ; --- URL
        url := NormalizeURL(path)
        if (url = "") {
            MsgBox("Not a valid http/https URL or existing file path."
                , "Add Reference", "Iconx")
            return
        }
        Prefs.References().Push(Map(
            "name", cleanName,
            "type", "url",
            "path", url,
            "uses", 0
        ))
    }

    Prefs.Save()
    g.Destroy()
}

; ---- manage / remove --------------------------------------------------------

ShowReferencesManager() {
    dark := Prefs.Get("display", "darkMode", false)
    palette := ModernPalette(dark)

    MouseGetPos(&mx, &my)
    work := GetWorkAreaAt(mx, my)
    w := 560, h := 360
    ; Same unit contract as ShowAddReferenceDialog: logical w/h, clamp
    ; using physical pixels.
    scale := A_ScreenDPI / 96
    pos := ClampToWorkArea(mx + 12, my + 12, Round(w * scale), Round(h * scale), work)

    g := Gui("+AlwaysOnTop -MaximizeBox -MinimizeBox", "Manage References")
    g.MarginX := 16, g.MarginY := 14
    g.BackColor := palette["bg"]
    g.SetFont("s10 c" palette["fg"], ModernFont())

    lv := g.Add("ListView"
        , "x16 y16 w528 h260 Grid +LV0x10000 Background" palette["bgAlt"] " c" palette["fg"]
        , ["Name","Type","Uses","Path"])

    for ref in Prefs.References() {
        if !(ref is Map) || !ref.Has("name")
            continue
        lv.Add(""
            , ref["name"]
            , ref.Has("type") ? ref["type"] : ""
            , ref.Has("uses") ? ref["uses"] : 0
            , ref.Has("path") ? ref["path"] : "")
    }
    lv.ModifyCol(1, 140)
    lv.ModifyCol(2, 60)
    lv.ModifyCol(3, 60)
    lv.ModifyCol(4, 260)

    btnRemove := g.Add("Button"
        , "x16 y290 w140 h32 +Background" palette["btnBg"] " c" palette["btnFg"]
        , "Remove Selected")
    btnRemove.OnEvent("Click", (*) => RemoveSelectedReferences(lv, g))

    btnAdd := g.Add("Button"
        , "x166 y290 w120 h32 +Background" palette["btnBg"] " c" palette["btnFg"]
        , "Add...")
    btnAdd.OnEvent("Click", (*) => (g.Destroy(), ShowAddReferenceDialog()))

    btnClose := g.Add("Button"
        , "Default x460 y290 w84 h32 +Background" palette["btnBg"] " c" palette["btnFg"]
        , "Close")
    btnClose.OnEvent("Click", (*) => g.Destroy())

    g.OnEvent("Close",  (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())

    ApplyModernChrome(g, dark)
    g.Show("x" pos.x " y" pos.y " w" w " h" h)
    WinActivate("ahk_id " g.Hwnd)
}

RemoveSelectedReferences(lv, g) {
    rows := []
    row := 0
    while (row := lv.GetNext(row))
        rows.Push(row)
    if (rows.Length = 0)
        return
    if (MsgBox("Remove " rows.Length " reference" (rows.Length > 1 ? "s" : "") "?"
            , "Manage References", "YesNo Iconi") != "Yes")
        return

    ; Collect names first (LV row indices change after deletion)
    names := []
    for r in rows
        names.Push(lv.GetText(r, 1))

    for name in names {
        idx := Prefs.FindReferenceIndex(name)
        if (idx != 0)
            Prefs.References().RemoveAt(idx)
    }
    Prefs.Save()
    g.Destroy()
    ShowReferencesManager()
}
