try {
    batch_convert();
} catch (e) {
    alert(e.message + " (line " + e.line + ")");
}

function batch_convert() {
    function find_documents(params) {
        function find_opendocs(docs, params) {
            var array = [];
            for (var i = 0; i < docs.length; i += 1) {
                try {
                    if (docs[i].saved == false) {
                        docs[i].save(File(params.output_folder + docs[i].name));
                    }
                    array.push(docs[i].fullName);
                } catch (e) {
                    batch_problems.push(docs[i].name + ": " + e.message);
                }
            }
            return array;
        }
        if (app.documents.length > 0) {
            return find_opendocs(app.documents, params);
        } else {
            return find_files(params.input_folder, params.include_subdir, params.source_type);
        }
    }
    var viewPdf = app.pdfExportPreferences.viewPDF;
    var params = get_data(File(script_dir() + "batch_convert.txt"));
    var batch_problems = [];
    var all_docs = find_documents(params);
    batch_problems = process_docs(all_docs, params, batch_problems);
    try {
        Window.find("palette", "Processing").close();
    } catch (_) {

    }
    app.pdfExportPreferences.viewPDF = viewPdf;
    if (batch_problems.length > 0) {
        alert("Problems with:\r\r" + batch_problems.join("\r"));
    }
}

function process_docs(docs, params, batch_problems) {
    var iaPdfView = app.interactivePDFExportPreferences.viewPDF;
    app.interactivePDFExportPreferences.viewPDF = false;

    function strip_drive(doc) {
        return String(doc.filePath).replace(/\/[^\/]+\//, "");
    }

    function getScriptLanguage(s) {
        var fileType = s.toLowerCase().replace(/.+\./, "").replace(/jsx?(bin)?/, "js");
        switch (fileType) {
            case "js":
                return ScriptLanguage.JAVASCRIPT;
            case "scpt":
                return ScriptLanguage.APPLESCRIPT_LANGUAGE;
            case "vbs":
                return ScriptLanguage.VISUAL_BASIC;
            default:
                return null;
        }
    }
    if (params.runscript_check) {
        params.script = File(params.selected_script);
        params.language = getScriptLanguage(params.selected_script);
    }
    if (params.source_name == "Word") {
        return convert_word(params);
    }
    if (params.target_type == "AEM") {
        app.pdfExportPreferences.viewPDF = false;
    }
    if (params.target_type == "PDF") {
        app.pdfExportPreferences.viewPDF = params.viewPDFstate;
    }
    var longest_name = get_longest(docs) + 6;
    var show_filename = create_message_window(longest_name);
    show_filename.show(); //show()
    var re = /\.[^\.]+$/;
    if (params.target_type == "JPG" && params.jpeg_preset != "[None]") {
        var jpeg_properties = get_jpeg_properties(params.jpeg_preset);
    } else if (params.target_type == "HTML" && params.html_preset != "[None]") {
        var html_properties = get_html_properties(params.html_preset);
    } else {
        if (params.target_type == "AEM" && params.AEM_preset != "[None]") {
            var AEM_properties = get_AEM_properties(params.AEM_preset);
        }
    }
    for (var i = 0; i < docs.length; i += 1) {
        try {
            current_doc = open_doc(docs[i], params);
            if (params.target_type == "AEM") {
                output_name = docs[i].name.replace(re, ".article");
            } else {
                output_name = docs[i].name.replace(re, "." + params.target_type.toLowerCase());
            }
            if (params.runscript_check) {
                app.doScript(params.script, params.language);
            }
            if (params.output_folder == "") {
                outfile = new File(File(docs[i].fullName).path + "/" + output_name);
            } else {
                outfile = new File(params.output_folder + output_name);
            }
            try {
                show_filename.message.text = decodeURI(docs[i].fullName) + " (" + Number(i + 1) + " of " + docs.length + ")";
            } catch (_) {
                show_filename.message.text = decodeURI(docs[i].fullName.replace(/%/g, "_")) + " (" + Number(i + 1) + " of " + docs.length + ")";
            }
            if (params.overwrite == false) {
                outfile = File(unique_name(outfile));
            }
            if (params.update_links) {
                try {
                    current_doc.links.everyItem().update();
                } catch (_) {
                    batch_problems.push("Missing links in " + decodeURI(docs[i]));
                }
            }
            switch (params.target_type) {
                case "PDF":
                    if (params.interactivePDF == false) {
                        if (params.separate_PDF_pages) {
                            exportPerPage(current_doc, params);
                        } else {
                            app.pdfExportPreferences.pageRange = PageRange.allPages;
                            export_PDF(current_doc, outfile, params);
                        }
                    } else {
                        current_doc.exportFile(ExportFormat.INTERACTIVE_PDF, outfile, false);
                    }
                    break;
                case "ICML":
                    exportICML(current_doc, outfile);
                    break;
                case "EPS":
                    if (params.outlines == true) {
                        var stories = current_doc.stories.everyItem().getElements();
                        for (var zz = 0; zz < stories.length; zz += 1) {
                            try {
                                stories[i].createOutlines(true);
                            } catch (_) {

                            }
                        }
                    }
                    current_doc.exportFile(ExportFormat.epsType, outfile, false);
                    break;
                case "HTML":
                    if (params.html_preset != "[None]") {
                        current_doc.htmlExportPreferences.properties = html_properties;
                    }
                    current_doc.htmlExportPreferences.viewDocumentAfterExport = params.view_html;
                    current_doc.exportFile(ExportFormat.html, outfile);
                    break;
                case "JPG":
                    if (params.jpeg_preset != "[None]") {
                        app.jpegExportPreferences.properties = jpeg_properties;
                    }
                    current_doc.exportFile(ExportFormat.JPG, outfile, false);
                    break;
                case "INX":
                    current_doc.exportFile(ExportFormat.indesignInterchange, outfile);
                    break;
                case "IDML":
                    current_doc.exportFile(ExportFormat.indesignMarkup, outfile);
                    break;
                case "XML":
                    current_doc.exportFile(ExportFormat.xml, outfile);
                    break;
                case "PNG":
                    current_doc.exportFile(ExportFormat.PNG_FORMAT, outfile, true);
                    break;
                case "EPUB":
                    //CountPages();    
                    createTOC();  
                    halfSpace();                                                              
                    setEPubExportOptions();
                    current_doc.exportFile(ExportFormat.EPUB, outfile, false);
                    break;
                case "SWF":
                    current_doc.exportFile(ExportFormat.SWF, outfile, false);
                    break;
                case "RTF":
                    rtf_story(current_doc).exportFile(ExportFormat.rtf, outfile);
                    break;
                case "INDD":
                    current_doc.save(outfile);
                    break;
                case "INDT":
                    current_doc.save(outfile);
                    break;
                case "AEM":
                    app.exportDpsArticle(outfile, current_doc, AEM_properties);
                case "PACK":
                    if (params.preserve_structure_for_package == false) {
                        outfolder = Folder(params.output_folder);
                    } else {
                        var outfolder = Folder(params.output_folder + current_doc.name.replace(/\.indd$/, ""));
                        if (outfolder.exists == false) {
                            outfolder.create();
                        }
                    }
                    if (parseInt(app.version) <= 9) {
                        current_doc.packageForPrint(outfolder, params.pack_fonts, params.pack_links, true, params.pack_updateGraphics, params.pack_hidden, true, params.create_report, "", false);
                    } else if (parseInt(app.version) <= 12) {
                        current_doc.packageForPrint(outfolder, params.pack_fonts, params.pack_links, true, params.pack_updateGraphics, params.pack_hidden, true, params.create_report, false, false, "", "", false);
                    } else {
                        if (parseInt(app.version) >= 13) {
                            current_doc.packageForPrint(outfolder, params.pack_fonts, params.pack_links, true, params.pack_updateGraphics, params.pack_hidden, true, params.create_report, false, false, "", params.docHyphenationOnly, "", false);
                        }
                    }
                    if (params.pack_pdf) {
                        app.pdfExportPreferences.pageRange = PageRange.allPages;
                        var outfile = File(outfolder + "/" + current_doc.name.replace(/\.indd$/i, ".pdf"));
                        export_PDF(current_doc, outfile, params);
                    }
                    if (params.pack_idml) {
                        var outfile = File(outfolder + "/" + current_doc.name.replace(/\.indd$/i, ".idml"));
                        current_doc.exportFile(ExportFormat.INDESIGN_MARKUP, outfile);
                    }
                    if (params.pack_jpg) {
                        var outfile = File(outfolder + "/" + current_doc.name.replace(/\.indd$/i, ".jpg"));
                        current_doc.exportFile(ExportFormat.JPG, outfile, false);
                    }
                    if (params.pack_png) {
                        var outfile = File(outfolder + "/" + current_doc.name.replace(/\.indd$/i, ".png"));
                        current_doc.exportFile(ExportFormat.PNG_FORMAT, outfile, false);
                    }
            }
        } catch (e) {
            batch_problems.push(decodeURI(docs[i]) + ": " + e.message + " (" + e.line + ")");
        }
        if (params.save_docs) {
            current_doc.save()
        }
        if (params.close_open_docs) {
            try {
                current_doc.close(SaveOptions.no);
            } catch (_) {
                batch_problems.push("Problem closing " + decodeURI(docs[i]));
            }
        }
    }
    if (params.close_open_docs) {
        for (var i = app.documents.length - 1; i > -1; i--) {
            app.documents[i].close(SaveOptions.no);
        }
    }
    app.interactivePDFExportPreferences.viewPDF = iaPdfView;
    return batch_problems;
}

function exportICML(doc, f) {
    var stories = doc.stories.everyItem().getElements();
    var baseName = String(doc.fullName).replace(/\.indd$/, "_");
    for (var i = stories.length - 1; i >= 0; i--) {
        if (stories[i].textContainers[0].parent instanceof Spread) {
            f = File(baseName + stories[i].id + ".icml");
            stories[i].exportFile(ExportFormat.INCOPY_MARKUP, f);
        }
    }
}

function exportPerPage(doc, params) {
    var base = doc.name.replace(/\.indd$/, "");
    if (params.pdf_per_page) {
        var pnames = doc.pages.everyItem().name;
    } else {
        var pnames = doc.pages.everyItem().documentOffset;
        for (var i = 0; i < pnames.length; i += 1) {
            pnames[i] = String(pnames[i] + 1);
        }
    }

    function pad(s) {
        if (!isNaN(s)) {
            return "00" + s.slice(-3);
        }
        return s;
    }
    for (var i = 0; i < pnames.length; i += 1) {
        var outfile = File(params.output_folder + base + "_" + pad(pnames[i]) + ".pdf");
        app.pdfExportPreferences.pageRange = pnames[i];
        export_PDF(doc, outfile, params);
    }
}

function export_PDF(current_doc, outfile, params) {
    for (var i = 0; i < params.pdf_preset.length; i += 1) {
        n = params.pdf_preset[i].indexOf("_") == 0 ? params.pdf_preset[i] : "";
        f = File(outfile.path + "/" + outfile.name.replace(/\.(?:ind[db]|pdf)$/i, n + ".pdf"));
        current_doc.exportFile(ExportFormat.PDF_TYPE, f, false, app.pdfExportPresets.item(params.pdf_preset[i]));
    }
}

function word_files(dir) {
    var list = Folder(dir).getFiles(function (f) {
        return f.name.search(/\.(?:rtf|docx?)$/gi) !== -1;
    });
    return list;
}

function getDocument(params) {
    var f = File(params.output_folder + "/Word_InDesign_Template.indd");
    if (f.exists) {
        return app.open(f);
    }
    f = File(params.output_folder + "/Word_InDesign_Template.idml");
    if (f.exists) {
        return app.open(f);
    }
    return app.documents.add();
}

function findStory(frames) {
    if (frames.length === 1) {
        return frames[0].parentStory;
    }
    for (var i = frames.length - 1; i >= 0; i--) {
        if (/main[_ ]?frame/i.test(frames[i].name)) {
            return frames[i].parentStory;
        }
    }
    return null;
}

function get_story(params) {
    var doc = getDocument(params);
    var story = findStory(doc.pages[0].textFrames);
    if (story !== null) {
        return story;
    }
    doc.zeroPoint = [0, 0];
    doc.textPreferences.smartTextReflow = true;
    try {
        doc.textPreferences.limitToMasterTextFrames = false;
    } catch (_) {

    }
    var pb = doc.pages[0].bounds;
    var m = doc.pages[0].marginPreferences;
    var gb = [m.top, m.left, pb[2] - m.bottom, pb[3] - m.left];
    doc.pages[0].textFrames.add({
        geometricBounds: gb
    });
    doc.pages.add();
    doc.pages[1].textFrames.add({
        geometricBounds: gb
    });
    doc.pages[0].textFrames[0].nextTextFrame = doc.pages[1].textFrames[0];
    return doc.pages[0].textFrames[0].parentStory;
}

function delete_empty_frames() {
    app.findGrepPreferences = app.changeGrepPreferences = null;
    app.findGrepPreferences.findWhat = "\\A\\Z";
    var empties = app.documents[0].findGrep(true);
    for (var i = 0; i < empties.length; i += 1) {
        empties[i].parentTextFrames[0].remove();
    }
}

function convert_word(params) {
    var problems = [];
    var w = word_files(params.input_folder);
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
    if (params.target_type == "InDesign single document") {
        placedstory = get_story(params);
        for (var i = 0; i < w.length; i += 1) {
            try {
                if (placedstory.contents.length > 0 && placedstory.characters[-1].contents !== "\r") {
                    placedstory.insertionPoints[-1].contents = "\r";
                }
                placedstory.insertionPoints[-1].contents = "&&&" + w[i].name + "&&&\r";
                placedstory.insertionPoints[-1].place(w[i]);
                if (params.runscript_check) {
                    app.doScript(params.script, params.language);
                }
            } catch (_) {
                problems.push("Problem placing " + w.name);
            }
        }
        delete_empty_frames();
        placedstory.recompose();
        app.documents[0].save(File(params.output_folder + w[0].name.replace(/\.(?:rtf|docx?)$/i, ".indd")));
        app.documents[0].close(SaveOptions.NO);
    } else {
        for (var i = 0; i < w.length; i += 1) {
            placedstory = get_story(params);
            try {
                placedstory.insertionPoints[-1].place(w[i]);
                if (params.runscript_check) {
                    app.doScript(params.script, params.language);
                }
            } catch (_) {
                problems.push("Problem placing " + w.name);
            }
            delete_empty_frames();
            app.documents[0].save(File(params.output_folder + w[i].name.replace(/\.(?:rtf|docx?)$/i, ".indd")));
            app.documents[0].close(SaveOptions.NO);
        }
    }
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
    return problems;
}

function open_doc(f, params) {
    if (params.ignore_errors) {
        app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
    } else {
        app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
    }
    if (app.documents.length > 0 || params.runscript_check || !params.ignore_errors) {
        app.open(f);
    } else {
        app.open(f, false);
    }
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
    return /\.indb$/.test(f.name) ? app.books[0] : app.documents[0];
}

function get_longest(docs) {
    var longest = 0;
    for (var i = 0; i < docs.length; i += 1) {
        longest = Math.max(longest, docs[i].fullName.length);
    }
    return longest;
}

function create_message_window(le) {
    var w = new Window("palette {text: \"Processing\"}");
    w.message = w.add("statictext", undefined, "");
    w.message.characters = le;
    return w;
}

function get_data(history) {
    var icons = define_icons();
    var w = new Window("dialog", "Batch process", undefined, {
        closeButton: false
    });
    w.orientation = "row";
    w.alignChildren = ["top", "top"];
    var w1 = w.add("group");
    w1.alignChildren = "fill";
    var main = w1.add("group {alignChildren: \"fill\", orientation: \"column\"}");
    var folder = main.add("panel {alignChildren: \"right\"}");
    var infolder = folder.add("group {_: StaticText {text: \"Input folder:\"}}");
    var inp = infolder.add("group {alignChildren: \"left\", orientation: \"stack\"}");
    if (File.fs != "Windows") {
        var inlist = inp.add("dropdownlist");
    }
    var infolder_name = inp.add("edittext");
    if (File.fs == "Windows") {
        var inlist = inp.add("dropdownlist");
    }
    inlist.preferredSize = [330, 22];
    infolder_name.preferredSize = [310, 22];
    var infolder_button = infolder.add("iconbutton", undefined, {
        style: "toolbutton"
    });
    var outfolder = folder.add("group {orientation: \"row\", _: StaticText {text: \"Output folder:\"}}");
    var outfolder_name = outfolder.add("edittext");
    outfolder_name.preferredSize.width = 330;
    var outfolder_button = outfolder.add("iconbutton", undefined, {
        style: "toolbutton"
    });
    var check_boxes = folder.add("group {orientation: \"row\"}");
    check_boxes.margins.right = 45;
    var subfolders = check_boxes.add("checkbox", undefined, "Include subfolders");
    subfolders.enabled = app.documents.length == 0;
    var ignore_errors = check_boxes.add("checkbox", undefined, "Ignore errors");
    var overwrite = check_boxes.add("checkbox", undefined, "Overwrite existing files");
    var formats = main.add("panel {orientation: \"row\"}");
    var source_group = formats.add("group");
    source_group.add("statictext", undefined, "Source format:");
    var source_list = source_group.add("dropdownlist", undefined, ["InDesign", "InDesign book", "InDesign (template)", "INX", "IDML", "PageMaker", "QuarkExpress", "Word"]);
    source_list.preferredSize.width = 130;
    formats.add("statictext", undefined, "Target format:");
    var target_list = formats.add("dropdownlist", undefined, ["InDesign", "InDesign (template)", "IDML", "ICML", "PDF", "PDF (Interactive)", "EPS", "RTF", "HTML", "XML", "JPG", "EPUB", "PNG", "SWF", "Package", "AEM (Adobe Experience Manager)"]);
    target_list.preferredSize.width = 130;
    var swapFormats = formats.add("button {text: \"X\"}");
    swapFormats.preferredSize.width = 25;
    var pdfOptions = main.add("panel {text: \"PDF\", alignChildren: [\"left\", \"top\"], orientation: \"row\"}");
    var pdf_presets = app.pdfExportPresets.everyItem().name;
    var pdf_presetGroup = pdfOptions.add("group {alignChildren: [\"left\", \"top\"]}");
    pdf_presetGroup.add("statictext {text: \"PDF preset:\"}");
    pdf_presetlist = pdf_presetGroup.add("listbox", undefined, pdf_presets, {
        multiselect: true
    });
    pdf_presetlist.preferredSize.height = 107;
    pdf_presetlist.preferredSize.width = 200;
    pdf_presetlist.selection = 0;
    var pdfOptions2 = pdfOptions.add("group {orientation: \"column\", alignChildren: \"left\"}");
    pdfOptions2.spacing = 2;
    var view_pdf = pdfOptions2.add("checkbox", undefined, "View PDFs after export");
    var separate_PDF_pages = pdfOptions2.add("checkbox {text: \"Export separate pages\"}");
    var dummy1 = pdfOptions2.add("group");
    dummy1.margins[0] = 17;
    var per_page = dummy1.add("panel {alignChildren: \"left\"}");
    per_page.spacing = 0;
    var pdf_per_page = per_page.add("radiobutton {text: \"Use page numbers\"}");
    var pdf_per_offset = per_page.add("radiobutton {text: \"Use document offsets\"}");
    var packaging_group = main.add("panel {text: \"Package\", alignChildren: \"left\"}");
    var include_groupA = packaging_group.add("group {orientation: \"row\"}");
    var pack_pdf = include_groupA.add("checkbox {text: \"Include PDF\"}");
    var pack_idml = include_groupA.add("checkbox {text: \"Include IDML\"}");
    var pack_jpg = include_groupA.add("checkbox {text: \"Include JPEG\"}");
    var pack_png = include_groupA.add("checkbox {text: \"Include PNG\"}");
    var pack_links = include_groupA.add("checkbox {text: \"Include links\"}");
    var include_groupB = packaging_group.add("group {orientation: \"row\"}");
    var pack_updateGraphics = include_groupB.add("checkbox {text: \"Update graphics\"}");
    var pack_fonts = include_groupB.add("checkbox {text: \"Include fonts\"}");
    var pack_hidden = include_groupB.add("checkbox {text: \"Include hidden and non-printing content\"}");
    var include_groupC = packaging_group.add("group {orientation: \"row\"}");
    var preserve_structure_for_package = include_groupC.add("checkbox {text: \"Preserve folder structure\"}");
    var docHyphenationOnly = include_groupC.add("checkbox {text: \"Doc hyphenation exc. only\"}");
    var create_report = include_groupC.add("checkbox {text: \"Create report\"}");
    var options = main.add("panel {text: \"Miscellaneous\", alignChildren: \"left\"}");
    var jpeg_presetGroup = options.add("group");
    jpeg_presetGroup.add("statictext {text: \"JPEG preset:\", characters: 8}");
    var jpeg_presetlist = jpeg_presetGroup.add("dropdownlist", undefined, find_jpeg_presets());
    jpeg_presetlist.preferredSize.width = 200;
    jpeg_presetlist.selection = 0;
    if (parseFloat(app.version) > 7) {
        var html_presetGroup = options.add("group");
        html_presetGroup.add("statictext {text: \"HTML preset:\", characters: 8}");
        var html_presetlist = html_presetGroup.add("dropdownlist", undefined, find_html_presets());
        html_presetlist.preferredSize.width = 200;
        html_presetlist.selection = 0;
        var view_html = html_presetGroup.add("checkbox", undefined, "View HTML after export");
    }
    if (parseFloat(app.version) >= 6) {
        var AEM_presetGroup = options.add("group");
        AEM_presetGroup.add("statictext {text: \"AEM preset:\", characters: 8}");
        var AEM_presetlist = AEM_presetGroup.add("dropdownlist", undefined, find_AEM_presets());
        AEM_presetlist.preferredSize.width = 200;
        AEM_presetlist.selection = AEM_presetlist.items.length > 1 ? 1 : 0;
    }
    var misc_options = options.add("group {alignChildren: \"left\", orientation: \"column\"}");
    var outlines = misc_options.add("checkbox", undefined, "Convert text to outlines (EPS export)");
    var update_links = misc_options.add("checkbox", undefined, "Update modified links before exporting (missing links are always ignored)");
    var close_open_docs = misc_options.add("checkbox", undefined, "Close open documents");
    var runscript = misc_options.add("group");
    var runscript_check = runscript.add("checkbox {text: \"Run a script:\"}");
    var runscriptScript = runscript.add("edittext {characters: 30}");
    var runscriptIcon = runscript.add("iconbutton", undefined, {
        style: "toolbutton"
    });
    var save_docs = misc_options.add("checkbox", undefined, "Save changed documents on closing");
    var bpresets = get_batch_presets();
    var batch_preset_panel = main.add("panel {orientation: \"row\"}");
    batch_preset_panel.add("statictext {text: \"Batch processor preset:\"}");
    var batch_presets = batch_preset_panel.add("dropdownlist", undefined, bpresets);
    var save_preset_icon = batch_preset_panel.add("iconbutton", undefined, {
        style: "toolbutton"
    });
    var delete_preset_icon = batch_preset_panel.add("iconbutton", undefined, {
        style: "toolbutton"
    });
    batch_presets.selection = 0;
    batch_presets.preferredSize.width = 220;
    var buttons = w.add("group {orientation: \"column\", alignChildren: \"fill\"}");
    var okButton = buttons.add("button", undefined, "OK");
    buttons.add("button", undefined, "Cancel", {
        name: "cancel"
    });
    helptip_button = buttons.add("checkbox {text: \"Show tool tips\"}");
    save_settings = buttons.add("checkbox {text: \"Save settings\"}");

    function find_jpeg_presets() {
        var presets = Folder(script_dir()).getFiles("*.jpeg_preset");
        var file_array = ["[None]"];
        for (var i = 0; i < presets.length; i += 1) {
            file_array.push(presets[i].name.replace(".jpeg_preset", ""));
        }
        return file_array;
    }

    function find_html_presets() {
        var presets = Folder(script_dir()).getFiles("*.html_preset");
        var file_array = ["[None]"];
        for (var i = 0; i < presets.length; i += 1) {
            file_array.push(presets[i].name.replace(".html_preset", ""));
        }
        return file_array;
    }

    function find_AEM_presets() {
        var presets = Folder(script_dir()).getFiles("*.AEM_preset");
        var file_array = ["[You must install presets]"];
        for (var i = 0; i < presets.length; i += 1) {
            file_array.push(presets[i].name.replace(".AEM_preset", ""));
        }
        return file_array;
    }

    function get_batch_presets() {
        var p = [];
        var f = Folder(script_dir()).getFiles("*.batch_preset");
        for (var i = 0; i < f.length; i += 1) {
            p.push(decodeURI(f[i].name.replace(/\.batch_preset$/, "")));
        }
        if (p.length > 1) {
            p.sort()
        }
        p.unshift("[None]");
        return p;
    }
    batch_presets.onChange = function () {
        if (batch_presets.selection.text == "[None]") {
            return;
        }
        var path = script_dir();
        var previous = get_previous(File(path + batch_presets.selection.text + ".batch_preset"));
        set_dialog(previous);
    };
    delete_preset_icon.onClick = function () {
        var path = script_dir();
        if (askYN("Delete " + batch_presets.selection.text + "?") == true) {
            try {
                File(path + "/" + batch_presets.selection.text + ".batch_preset").remove();
                batch_presets.remove(batch_presets.find(batch_presets.selection.text));
                batch_presets.selection = 0;
            } catch (_) {
                alert("Could not remove " + batch_presets.selection.text);
            }
        }
    };
    save_preset_icon.onClick = function () {
        var outfile = get_filename(batch_presets.selection.text);
        if (outfile == "") {
            return;
        }
        if (batch_presets.find(outfile.menu_name) == null) {
            insert_item(batch_presets, outfile.menu_name);
        }
        batch_presets.selection = batch_presets.find(outfile.menu_name);
        var dlg_data = collect_dlg_data();
        store_settings(outfile.file, dlg_data);
    };

    function get_filename(str) {
        var path = script_dir();
        var name = get_name(str);
        if (name == "") {
            return "";
        }
        var f = File(path + name + ".batch_preset");
        while (f.exists && askYN("Preset exists -- replace?") == false) {
            name = get_name(str);
            if (name == "") {
                return "";
            }
            f = File(path + name + ".batch_preset");
        }
        return {
            file: f,
            menu_name: name
        };
    }

    function get_name(str) {
        var w = new Window("dialog {alignChildren: \"right\"}");
        var gr = w.add("group {_: StaticText {text: \"Save preset as:\"}}");
        var e = gr.add("edittext {characters: 20, active: true}");
        e.text = str;
        var buttons = w.add("group");
        buttons.add("button {text: \"Cancel\"}");
        var ok = buttons.add("button {text: \"OK\"}");
        if (w.show() == 1) {
            return e.text;
        }
        return "";
    }

    function askYN(s) {
        var w = new Window("dialog", "", undefined, {
            closeButton: false
        });
        w.add("group {_: StaticText {text: \"" + s + "\"}}");
        var buttons = w.add("group");
        buttons.add("button", undefined, "No", {
            name: "cancel"
        });
        buttons.add("button", undefined, "Yes", {
            name: "ok"
        });
        return w.show() == 1 ? true : false;
    }

    function insert_item(list_obj, new_item) {
        if (list_obj.find(new_item) == null) {
            var stop = list_obj.items.length;
            var i = 0;
            while (i < stop && new_item > list_obj.items[i].text) {
                i++;
            }
            list_obj.add("item", new_item, i);
        }
    }
    helptip_button.onClick = function () {
        set_helptips(helptip_button.value);
    };

    function set_helptips(val) {
        if (val) {
            outfolder_name.helpTip = "With \"Include subfolders\" checked and \"Output folder\" blank, documents are saved in their originating folders.\r\rWith \"Include subfolders\" checked and \"Output folder\" specified, documents from all subfolders are saved in the specified output folder.\r\rExisting files are overwritten.";
            ignore_errors.helpTip = "Errors relating to missing fonts and missing and/or modified links are ignored.";
            source_list.helpTip = parseInt(app.version) <= 8 ? "PageMaker: versions 6.0-7.0 only.\rQuarkExpress: versions 3.3-4.1x only." : "QuarkExpress: versions 3.3-4.1x only.";
            source_list.helpTip += "\rWord: Any rtf, doc and docx document. In the Target dropdown, choose to place all Word files in a single document or in separate Indesign documents.";
            target_list.helpTip = "CS3 reads and writes INX only;\rCS4 reads and writes INX and IDML;\rCS5+ read INX and IDML, write IDML only.\r\rPackaging: CS5 and later.\rHTML: CS5.5 and later.\rPNG, AEM: CS6 and later.\rSWF: CC and later.";
            swapFormats.helpTip = "Swap source and target formats.";
            pdf_presetlist.helpTip = "Select any number of presets. If a preset name starts with _ (underscore character), the name is added to the PDF file's name.";
            pack_pdf.helpTip = "Include a PDF of each InDesign file in the package.";
            pack_idml.helpTip = "Include an IDML of each InDesign file in the package.";
            separate_PDF_pages.helpTip = "Enable the export of document pages as seperate PDF files, using as sequential numbers either of the two options in the panel below.";
            pdf_per_page.helpTip = "Export InDesign pages as separate PDF files, using the printed page numbers as file-name suffixes.";
            pdf_per_offset.helpTip = "Export InDesign pages as separate PDF files, using the document offsets of the pages as file-name suffixes.";
            pack_updateGraphics.helpTip = "Update graphics before packaging.";
            preserve_structure_for_package.helpTip = "When UNCHECKED, all documents are packaged together in the output folder (all document fonts together, all links together, each in their own subfolders).\rWhen CHECKED, the folder structure is preserved, i.e. is recreated under the output folder.";
            update_links.helpTip = "Force links to update (even if \"Ignore errors\" is checked). Note that only modified links are handled; the script cannot deal with missing links.";
            close_open_docs.helpTip = "If the script is run when some documents are open, choose to close those documents after they're processed.";
            runscript_check.helpTip = "Run a script against documents before converting them or saving them back as InDesign files.\rTo save the changes a script made, check \"Save documents on closing\".";
            save_docs.helpTip = "Choose to save documents if they've changed (by running a script against them or if you opted to update any links).";
            if (parseFloat(app.version) > 7) {
                html_presetlist.helpTip = "To export HTML using each document's own export settings, select [None].";
            }
            if (parseFloat(app.version) >= 8) {
                AEM_presetlist.helpTip = "For AEM (previously DPS) export you must select a preset.\rSee www.kahrel.plus.com/batch_convert.html for details";
            }
            helptip_button.helpTip = "Disable tooltips";
        } else {
            outfolder_name.helpTip = "";
            ignore_errors.helpTip = "";
            source_list.helpTip = "";
            target_list.helpTip = "";
            swapFormats.helpTip = "";
            pdf_presetlist.helpTip = "";
            separate_PDF_pages.helpTip = "";
            pdf_per_page.helpTip = "";
            pdf_per_offset.helpTip = "";
            pack_pdf.helpTip = "";
            pack_idml.helpTip = "";
            pack_idml.pack_updateGraphics = "";
            preserve_structure_for_package.helpTip = "";
            update_links.helpTip = "";
            close_open_docs.helpTip = "";
            runscript_check.helpTip = "";
            save_docs.helpTip = "";
            if (parseFloat(app.version) > 7) {
                html_presetlist.helpTip = "";
            }
            if (parseFloat(app.version) >= 7) {
                AEM_presetlist.helpTip = "";
            }
            helptip_button.helpTip = "Enable tooltips";
        }
    }

    function find_scripts() {
        var script_files = Folder(script_dir()).getFiles(function (f) {
            return f.name.search(/\.scpt|\.jsx?(?:bin)?|\.vbs/i) > 0;
        });
        var file_array = [];
        var le = script_files.length;
        for (var i = 0; i < le; i += 1) {
            file_array[i] = script_files[i].name;
        }
        file_array.unshift("[None]");
        return file_array;
    }
    pack_pdf.onClick = function () {
        pdfOptions.enabled = this.value;
    };
    inlist.onChange = function () {
        infolder_name.text = outfolder_name.text = inlist.selection.text;
    };
    runscriptIcon.onClick = function () {
        var f = Folder(script_dir()).openDlg("Select script", "InDesign scripts:*.jsx;*.jsxbin;*.vbs;*.scpt", false);
        if (f === null) {
            runscriptScript.text = "";
        } else {
            runscriptScript.text = f;
        }
    };
    infolder_name.onChange = function () {
        this.text = this.text.replace(/([^\/])$/, "$1/");
        if (Folder(this.text).exists == false) {
            this.text = "Folder does not exist".toUpperCase();
        } else {
            outfolder_name.text = this.text;
        }
        this.active = true;
    };
    outfolder_name.onChange = function () {
        if (this.text != "" && Folder(this.text).exists == false) {
            this.text = "Folder does not exist";
        } else if (this.text != "" && Folder(this.text).exists == true) {
            this.text = this.text.replace(/([^\/])$/, "$1/");
        } else {
            this.active = true;
        }
    };
    infolder_button.onClick = function () {
        var f = Folder(infolder_name.text).selectDlg("Choose a folder");
        if (f != null) {
            infolder_name.text = f.fullName + "/";
            outfolder_name.text = f.fullName + "/";
            infolder_name.active = true;
        } else {
            return 0;
        }
    };
    outfolder_button.onClick = function () {
        var f = Folder(outfolder_name.text).selectDlg("Choose a folder");
        if (f != null) {
            outfolder_name.text = f.fullName + "/";
            outfolder_name.active = true;
        } else {
            return 0;
        }
    };
    source_list.onChange = function () {
        update_links.enabled = source_list.selection.text.indexOf("InDesign") > -1;
        swapFormats.enabled = getSwapState();
        if (source_list.selection.text === "Word") {
            if (target_list.find("InDesign single document") === null) {
                target_list.add("item", "InDesign single document", 1);
            }
            target_list.selection = target_list.find("InDesign");
        } else {
            target_list.remove("InDesign single document");
        }
        if (source_list.selection.text === "InDesign book") {
            target_list.selection = target_list.find("PDF");
            separate_PDF_pages.enabled = false;
            options.enabled = false;
        }
        if (source_list.selection.text !== "InDesign book") {
            options.enabled = true;
        }
        if (source_list.selection.text === "InDesign" && target_list.selection.text === "PDF") {
            separate_PDF_pages.enabled = true;
        }
    };
    target_list.onChange = function () {
        pdfOptions.enabled = target_list.selection.text == "PDF" || target_list.selection.text == "Package" && pack_pdf.value == true;
        okButton.enabled = true;
        if (target_list.selection.text == "AEM (Adobe Experience Manager)") {
            okButton.enabled = AEM_presetlist.items.length > 1;
        }
        jpeg_presetGroup.enabled = false;
        if (parseFloat(app.version) > 7) {
            html_presetGroup.enabled = false;
        }
        if (parseFloat(app.version) >= 6) {
            AEM_presetGroup.enabled = false;
        }
        outlines.value = false;
        outlines.enabled = false;
        packaging_group.enabled = target_list.selection.text == "Package";
        update_links.enabled = true;
        switch (target_list.selection.text) {
            case "AEM (Adobe Experience Manager)":
                AEM_presetGroup.enabled = true;
                break;
            case "HTML":
                html_presetGroup.enabled = true;
                break;
            case "JPG":
                jpeg_presetGroup.enabled = true;
                break;
            case "EPS":
                outlines.enabled = true;
                break;
        }
        swapFormats.enabled = getSwapState();
    };
    runscript_check.onClick = function () {
        runscriptScript.enabled = runscript_check.value;
        runscriptIcon.enabled = runscript_check.value;
    };
    swapFormats.onClick = function () {
        var target = target_list.selection.text;
        var source = source_list.selection.text;
        target_list.selection = target_list.find(source);
        source_list.selection = source_list.find(target);
    };
    separate_PDF_pages.onClick = function () {
        per_page.enabled = separate_PDF_pages.value;
    };
    var ID = parseFloat(app.version);
    if (ID < 6) {
        source_list.remove(source_list.find("IDML"));
        target_list.remove(target_list.find("IDML"));
        target_list.remove(target_list.find("Package"));
        target_list.add("item", "INX");
    }
    if (ID === 6) {
        target_list.add("item", "INX", 3);
    }
    if (ID < 7) {
        target_list.remove(target_list.find("PDF (Interactive)"));
    }
    if (ID < 7.5) {
        target_list.remove(target_list.find("HTML"));
    }
    if (ID < 8) {
        target_list.remove(target_list.find("PNG"));
    }
    if (ID < 9) {
        target_list.remove(target_list.find("SWF"));
    }
    if (ID > 8.9) {
        source_list.remove(source_list.find("PageMaker"));
    }
    if (history.exists == true) {
        var previous = get_previous(history);
        set_dialog(previous);
        if (previous.helptips == undefined) {
            helptip_button.value = true;
        } else {
            helptip_button.value = previous.helptips;
        }
    } else {
        source_list.selection = source_list.find("InDesign");
        target_list.selection = target_list.find("PDF");
        outlines.enabled = false;
        runscriptScript.text = "";
        if (parseInt(app.version) > 7) {
            view_html.value = false;
        }
    }

    function set_dialog(previous) {
        infolder_name.text = fill_list(inlist, previous.input_folder);
        outfolder_name.text = previous.output_folder;
        subfolders.value = previous.include_subdir;
        ignore_errors.value = previous.ignore_errors;
        overwrite.value = previous.overwrite || false;
        source_list.selection = source_list.find(previous.source_name);
        target_list.selection = target_list.find(previous.target_name);
        view_pdf.value = previous.viewPDFstate || false;
        separate_PDF_pages.value = previous.separate_PDF_pages || false;
        pdf_per_page.value = previous.pdf_per_page || false;
        pdf_per_offset.value = previous.pdf_per_offset || false;
        if (previous.pdf_preset instanceof Array) {
            pdf_presetlist.selection = null;
            for (var i = 0; i < previous.pdf_preset.length; i += 1) {
                if (pdf_presetlist.find(previous.pdf_preset[i]) !== null) {
                    pdf_presetlist.selection = pdf_presetlist.find(previous.pdf_preset[i]);
                }
            }
        } else {
            pdf_presetlist.selection = pdf_presetlist.find(previous.pdf_preset);
        }
        pack_pdf.value = previous.pack_pdf || false;
        pack_idml.value = previous.pack_idml || false;
        pack_jpg.value = previous.pack_jpg || false;
        pack_png.value = previous.pack_png || false;
        pack_updateGraphics.value = previous.pack_updateGraphics || false;
        pack_links.value = previous.pack_links || false;
        pack_fonts.value = previous.pack_fonts || false;
        pack_hidden.value = previous.pack_hidden || false;
        preserve_structure_for_package.value = previous.preserve_structure_for_package || false;
        docHyphenationOnly.value = previous.docHyphenationOnly || false;
        create_report.value = previous.create_report || false;
        if (previous.jpeg_preset) {
            jpeg_presetlist.selection = jpeg_presetlist.find(previous.jpeg_preset);
        } else {
            jpeg_presetlist.selection = 0;
        }
        if (parseFloat(app.version) > 7 && "html_preset" in previous) {
            html_presetlist.selection = html_presetlist.find(previous.html_preset);
            view_html.value = previous.view_html;
        }
        update_links.value = previous.update_links;
        outlines.value = previous.outlines;
        if (app.documents.length > 0) {
            infolder_name.text = "";
        }
        runscript_check.value = previous.runscript_check || false;
        runscriptScript.enabled = previous.runscript_check || false;
        runscriptScript.text = previous.selected_script || false;
        save_docs.value = previous.save_docs;
        if (previous.batch_preset != null) {
            batch_presets.selection = batch_presets.find(previous.batch_preset) || 0;
        }
    }
    set_helptips(helptip_button.value);
    if (app.documents.length > 0) {
        subfolders.value = false;
        infolder.enabled = false;
        source_list.selection = source_list.find("InDesign");
        source_group.enabled = false;
        outfolder_name.active = true;
    } else {
        close_open_docs.value = true;
        close_open_docs.enabled = false;
    }

    function getSwapState() {
        return "InDesign INX IDML".indexOf(source_list.selection.text) > -1 && "InDesign INX IDML".indexOf(target_list.selection.text) > -1 && source_list.selection.text != target_list.selection.text;
    }
    swapFormats.enabled = getSwapState();
    w.onShow = function () {
        save_settings.value = true;
        pdf_presetlist.revealItem(previous.pdf_preset[0]);
        separate_PDF_pages.notify();
        separate_PDF_pages.notify();
    };
    if (w.show() == 2) { //show()
        w.close();
        exit();
    } else {
        if (infolder_name.text == "FOLDER DOES NOT EXIST") {
            exit();
        }
        if (infolder_name.text == outfolder_name.text && target_list.selection.text == "Package" && preserve_structure_for_package == false) {
            alert("Package output folder cannot be the same as the input folder.");
            exit();
        }
        var dlg_data = collect_dlg_data();
        dlg_data.helptips = helptip_button.value;
        if (save_settings.value) {
            store_settings(history, dlg_data);
            store_settings(File(script_dir() + batch_presets.selection.text.replace(/[\[\]]/g, "") + ".batch_preset"), dlg_data);
        }
        w.close();
        dlg_data.input_folder = infolder_name.text;
        return dlg_data;
    }

    function presetName(w) {

    }

    function getPresetNames(listItemArray) {
        var list = [];
        for (var i = 0; i < listItemArray.length; i += 1) {
            list[i] = listItemArray[i].text;
        }
        return list;
    }

    function collect_dlg_data() {
        var obj = {
            input_folder: create_string(inlist, infolder_name.text),
            output_folder: outfolder_name.text,
            source_type: get_source_extension(source_list.selection.text),
            source_name: source_list.selection.text,
            target_type: get_target_extension(target_list.selection.text),
            target_name: target_list.selection.text,
            overwrite: overwrite.value,
            include_subdir: subfolders.value,
            ignore_errors: ignore_errors.value,
            separate_PDF_pages: separate_PDF_pages.value,
            pdf_per_page: pdf_per_page.value,
            pdf_per_offset: pdf_per_offset.value,
            pdf_preset: getPresetNames(pdf_presetlist.selection),
            pack_pdf: pack_pdf.value,
            pack_idml: pack_idml.value,
            pack_png: pack_png.value,
            pack_jpg: pack_jpg.value,
            pack_links: pack_links.value,
            pack_fonts: pack_fonts.value,
            pack_hidden: pack_hidden.value,
            pack_updateGraphics: pack_updateGraphics.value,
            preserve_structure_for_package: preserve_structure_for_package.value,
            docHyphenationOnly: docHyphenationOnly.value,
            create_report: create_report.value,
            outlines: outlines.value,
            update_links: update_links.value,
            save_docs: save_docs.value,
            close_open_docs: close_open_docs.value,
            viewPDFstate: view_pdf.value,
            runscript_enabled: runscript_check.value,
            runscript_check: runscript_check.value && runscriptScript.text !== "",
            selected_script: runscriptScript.text,
            interactivePDF: target_list.selection.text == "PDF (Interactive)",
            batch_preset: batch_presets.selection.text,
            jpeg_preset: jpeg_presetlist.selection.text
        };
        if (parseFloat(app.version) >= 8) {
            obj.AEM_preset = AEM_presetlist.selection.text;
        }
        if (parseFloat(app.version) > 7) {
            obj.html_preset = html_presetlist.selection.text;
            obj.view_html = view_html.value;
        }
        return obj;
    }

    function fill_list(list, str) {
        for (var i = list.items.length - 1; i > -1; i--) {
            list.remove(list.items[i]);
        }
        var array = str.split("£$£");
        for (var i = 0; i < array.length; i += 1) {
            list.add("item", array[i]);
        }
        return array[0];
    }

    function create_string(list, new_mask) {
        if (parseInt(app.version) == 6) {
            return new_mask;
        }
        list.remove(list.find(new_mask));
        if (list.items.length > 0) {
            list.add("item", new_mask, 0);
        } else {
            list.add("item", new_mask);
        }
        var str = "";
        var stop = Math.min(list.items.length, 8) - 1;
        for (var i = 0; i < stop; i += 1) {
            str += list.items[i].text + "£$£";
        }
        str += list.items[i].text;
        return str;
    }

    function get_source_extension(s) {
        switch (s) {
            case "InDesign":
                return ["INDD"];
            case "InDesign book":
                return ["INDB"];
            case "InDesign (template)":
                return ["INDT"];
            case "IDML":
                return ["IDML"];
            case "INX":
                return ["INX"];
            case "PageMaker":
                return ["PMD", "PM6", "P65"];
            case "QuarkExpress":
                return ["QXD"];
            case "Word":
                return ["RTF", "DOC", "DOCX"];
        }
    }

    function get_target_extension(s) {
        switch (s) {
            case "InDesign":
                return "INDD";
            case "InDesign (template)":
                return "INDT";
            case "Package":
                return "PACK";
            case "PDF (Interactive)":
                return "PDF";
            case "AEM (Adobe Experience Manager)":
                return "AEM";
            default:
                return s;
        }
    }

    function array_index(array) {
        for (var i = 0; i < array.length; i += 1) {
            if (array[i].value == true) {
                return i;
            }
        }
    }
}

function get_previous(f) {
    var temp = {};
    if (f.exists) {
        f.open("r");
        temp = f.read();
        f.close();
    }
    return eval(temp);
}

function store_settings(f, obj) {
    f.open("w");
    f.write(obj.toSource());
    f.close();
}

function script_dir() {
    try {
        return File(app.activeScript).path + "/";
    } catch (e) {
        return File(e.fileName).path + "/";
    }
}

function array_item(s, array) {
    for (var i = 0; i < array.length; i += 1) {
        if (s == array[i]) {
            return i;
        }
    }
    return 0;
}

function find_files(dir, incl_sub, mask_array) {
    var arr = [];
    for (var i = 0; i < mask_array.length; i += 1) {
        if (incl_sub == true) {
            arr = arr.concat(find_files_sub(dir, [], mask_array[i]));
        } else {
            arr = arr.concat(Folder(dir).getFiles("*" + mask_array[i]));
        }
    }
    return arr;
}

function find_files_sub(dir, array, mask) {
    var f = Folder(dir).getFiles("*.*");
    for (var i = 0; i < f.length; i += 1) {
        if (f[i] instanceof Folder && f[i].name[0] != ".") {
            find_files_sub(f[i], array, mask);
        } else {
            if (f[i].name[0] != "." && f[i].name.substr(-mask.length).toUpperCase() == mask) {
                array.push(f[i]);
            }
        }
    }
    return array;
}

function unique_name(f) {
    function strip_base(s) {
        return s.replace(/_\d+$/, "");
    }
    var str = String(f);
    var pos = str.lastIndexOf(".");
    var base = str.slice(0, pos);
    var type = str.slice(pos, str.length);
    var n = 0;
    while (File(base + type).exists) {
        base = strip_base(base) + "_" + String(++n);
    }
    return base + type;
}

function get_props(preset, extension) {
    var f = File(script_dir() + "/" + preset + extension);
    f.open("r");
    var p = f.read().split(/[\n\r]+/);
    f.close();
    return p;
}

function check_prop(parts) {
    switch (parts[0]) {
        case "externalStyleSheets":
            return parts[1].split(",");
        case "javascripts":
            return parts[1].split(",");
    }
    return parts[1];
}

function get_html_properties(preset) {
    var o = {};
    var props = get_props(preset, ".html_preset");
    for (var i = 0; i < props.length; i += 1) {
        parts = props[i].split(": ");
        if (parts[0] != "imageExtension" && parts[0] != "serverPath") {
            o[parts[0]] = eval(check_prop(parts));
        }
    }
    return o;
}

function get_jpeg_properties(preset) {
    var o = {};
    var props = get_props(preset, ".jpeg_preset");
    for (var i = 0; i < props.length; i += 1) {
        parts = props[i].split(": ");
        o[parts[0]] = eval(parts[1]);
    }
    return o;
}

function get_AEM_properties(preset) {
    var f = File(script_dir() + "/" + preset + ".AEM_preset");
    var options = [];
    f.open("r");
    var s = f.read();
    f.close();
    var options = s.split(/[\n\r]/);
    for (var i = 0; i < options.length; i += 1) {
        parts = options[i].split(/: ?/);
        if (!isNaN(parts[1])) {
            parts[1] = Number(parts[1]);
        } else {
            if (/true|false/.test(parts[1])) {
                parts[1] = eval(parts[1]);
            }
        }
        options[i] = parts;
    }
    return options;
}

function rtf_story(doc) {
    doc.masterSpreads.everyItem().pageItems.everyItem().locked = false;
    doc.masterSpreads.everyItem().pageItems.everyItem().remove();
    while (doc.groups.length > 0) {
        doc.groups.everyItem().ungroup();
    }
    inlines(doc);
    if (parseInt(app.version) > 6) {
        var rtf_frame = doc.masterSpreads[0].textFrames.add({
            name: "rtf",
            geometricBounds: ["2cm", "2cm", "15cm", "15cm"]
        });
    } else {
        var rtf_frame = doc.masterSpreads[0].textFrames.add({
            label: "rtf",
            geometricBounds: ["2cm", "2cm", "15cm", "15cm"]
        });
    }
    if (doc.articles.length) {
        var a = doc.articles.everyItem().getElements();
        for (var i = 0; i < a.length; i += 1) {
            if (a[i].articleMembers.length && a[i].articleMembers[0].itemRef.isValid && a[i].articleMembers[0].itemRef.parentStory.isValid) {
                try {
                    move_story(a[i].articleMembers[0].itemRef.parentStory, rtf_frame);
                } catch (_) {

                }
            }
        }
    } else {
        move_story(longest_story(doc), rtf_frame);
        if (doc.stories.length > 1) {
            for (var i = 0; i < doc.pages.length; i += 1) {
                while (doc.pages[i].textFrames.length > 0) {
                    move_story(doc.pages[i].textFrames[0].parentStory, rtf_frame);
                }
            }
        }
    }
    return rtf_frame.parentStory;
}

function move_story(story, target_frame) {
    var story_frames = story.textContainers;
    story.move(LocationOptions.AT_END, target_frame.parentStory);
    target_frame.parentStory.insertionPoints[-1].contents = "\r";
    for (var i = story_frames.length - 1; i > -1; i--) {
        story_frames[i].locked = false;
        story_frames[i].remove();
    }
}

function inlines(doc) {
    var st = doc.stories;
    for (var i = doc.stories.length - 1; i > -1; i--) {
        while (st[i].textFrames.length > 0) {
            ix = st[i].textFrames[-1].parent.index;
            st[i].textFrames[-1].texts[0].move(LocationOptions.after, st[i].insertionPoints[ix]);
            st[i].textFrames[-1].locked = false;
            st[i].textFrames[-1].remove();
        }
    }
}

function longest_story(doc) {
    var temp = doc.stories[0];
    var len = doc.stories[0].contents.length;
    if (doc.stories.length > 1) {
        for (var i = 1; i < doc.stories.length; i += 1) {
            if (doc.stories[i].contents.length > len) {
                len = doc.stories[i].contents.length;
                temp = doc.stories[i];
            }
        }
    }
    return temp;
}

function define_icons() {
    var o = {
        folder: "PNG\r\n\n\rIHDR_%.-\tpHYsgAMA±|ûQ cHRMz%ùÿéu0ê`:o_ÅFÞIDATxÚbüÿÿ?-@11ÐÍ CFRLÏ*eþ÷!yÿ÷ïÏßþÆdOz²ÈýÅ@¡¸f(##C­k9_iÔ<fÆ\tÙÒ1@%Ì@ÌHÁÄ<gÊF212,ÓqÍbàRf`ãf`ådøÿç'ÃþÙaÿþûqóñÏ%Ö¾þîbA¶d¨¡_Ø@¡¬ìüÌ<\\^¥§Pýöî.ÃÝ3¾ÿ<ÿÈÝÄ\r AÌP6!£sXYþÃÖ¡ô¿`b}¼©+Uáù¿ºýë_]Å@± 7Ì¥÷÷2È2p\n(¥þýùÅð÷÷o?¿1üù¡ÿþù\rÆ_ÆÃýI@^@üJÄîb¡¿>¿døüê:§ \"Ã¿¿¿þÿý4øØ¿Al ýï/PîXþÀ±+^ý<ÔÎ`Â\nIÛÊìÎÊÀÄôþ¯_`A]ùj8þñé)ÃÖ/î¿aØÂÎÆÎ\n5ø@adW7¶0üûÿKXèÒß`þ]qÝ¿¿0þï_ÃG1Üyþ}9ÐÐOÈI 0~xr 5ÃÿPW\rYqé_ ÿ\t²û20³pÌ¦\nþ<Â0OG®@Í À^bPäýþè«S'2Ü{þ};;Ç Ö@üq ÅP6&öiÛ^¾[NJ¿ÿüc&ß¿!l`Ù0þÌXßþe`eã\t5ð'9çù¡4+9÷ÔÐ/Püj8@!»ø?Tð'ÌCÈàïPúL XÐýÚüÈ\"æßÈÁ0²{³#oIEND®B`"
    };
    if (parseInt(app.version) < 9) {
        o.bin = "PNG\r\n\n\rIHDR* \tpHYsttÞfx\nOiCCPPhotoshop ICC profilexÚSgTSé=÷ÞôBKKoR RB&*!\tJ!¡ÙQÁEEÈ Q,\nØä!¢£Êûá{£kÖ¼÷æÍþµ×>ç¬ó³ÏÀH3Q5©BàÇÄÆáä.@\n$p³d!sý#ø~<<+\"À¾xÓÀMÀ0ÿêB\\Àt8K@zB¦@F&S `ËcbãP-`'æÓø{[!  eDh;¬ÏVEX0fKÄ9Ø-0IWfH°·ÀÎ²0Q){`È##xFòW<ñ+®ç*x²<¹$9E[-qWW.(ÎI+6aa@.Ây24àóÌ àóýxÎ®ÎÎ6¶_-ê¿ÿ\"bbãþåÏ«p@át~Ñþ,/³;mþ¢%îh^ u÷f²@µ éÚWópø~<<E¡¹ÙÙåääØJÄB[aÊW}þgÂ_ÀWýlù~<ü÷õà¾â$2]GøàÂÌôL¥Ï\tbÜæGü·ÿüÓ\"ÄIb¹X*ãQqDó2¥\"B)Å%Òÿdâß,û>ß5°j>{-¨]cöK'XtÀâ÷ò»oÁÔ(háÏwÿï?ýG %fIq^D$.TÊ³?ÇD *°AôÁ,ÀÁÜÁü`6B$ÄÂBB\ndr`)¬B(Í°*`/Ô@4ÀQhp.ÂU¸=púaÁ(¼\tAÈa!ÚbX#ø!ÁH$ ÉQ\"K5H1RT UHò=r9\\Fº;È2ü¼G1²Q=ÔµC¹¨7F¢Ðdt1 Ðr´=6¡çÐ«hÚ>CÇ0Àè3Äl0.ÆÃB±8,\tcË±\"¬«Æ°V¬»õcÏ±wEÀ\t6wB aAHXLXNØH¨ $4Ú\t7\tQÂ'\"¨K´&ºùÄb21XH,#Ö/{CÄ7$C2'¹I±¤TÒÒFÒnR#é,©4H#ÉÚdk²9, +ÈääÃä3ää!ò[\nb@q¤øSâ(RÊjJåå4åe2AU£RÝ¨¡T5ZB­¡¶R¯Q¨4u9ÍIK¥­¢Óhh÷i¯ètºÝNÐWÒËéGèèôw\rÇg(gw¯L¦ÓÇT071ëçoUX*¶*|Ê\nJ&*/T©ª¦ªÞªUóUËT©^S}®FU3Sã©\tÔ«UªPëSSg©;¨ªg¨oT?¤~YýYÃLÃOC¤Q ±_ã¼Æ c³x,!k\r«u5Ä&±ÍÙ|v*»ý»=ª©¡9C3J3W³Róf?ãqøtN\tç(§ó~Þï)â)¦4L¹1e\\kªX«H«Q«Gë½6®í§¦½E»YûAÇJ'\\'GgÎçSÙSÝ§\n§M=:õ®.ªk¥¡»Dw¿n§î¾^Lo§Þy½çú}/ýTýmú§õGX³$ÛÎ<Å5qo</ÇÛñQC]Ã@C¥aaá¹Ñ<£ÕFFiÆ\\ã$ãmÆmÆ£&&!&KMêMîRM¹¦)¦;L;LÇÍÌÍ¢ÍÖ5=1×2çç×ß·`ZxZ,¶¨¶¸eI²äZ¦Yî¶¼nZ9Y¥XUZ]³F­­%Ö»­»§§¹NN«ÖgÃ°ñ¶É¶©·°åØÛ®¶m¶}agbg·Å®Ãî½}º}ý=\rÙ«Z~s´r:V:ÞÎî?}Åôé/gXÏÏØ3ã¶Ë)ÄiSÓGgg¹sóKË.>.ÆÝÈ½äJtõq]ázÒõ³Âí¨Û¯î6îiîÜÌ4)Y3sÐÃÈCàQåÑ?0kß¬~OCOgµç#/c/W­×°·¥wª÷aï>ö>rã>ã<7Þ2ÞY_Ì7À·È·ËOÃo_ßC#ÿdÿzÿÑ§%gA[ûøz|!¿?:Ûeö²ÙíA ¹AA­åÁ­!hÈì­!÷çÎÎiP~èÖÐaæaÃ~'W?pXÑ15wÑÜCsßDúDDÞg1O9¯-J5*>ª.j<Ú7º4º?Æ.fYÌÕXXIlK9.*®6nl¾ßüíóââã{/È]py¡ÎÂô§©.,:@LN8ðA*¨%òw%\nyÂÂg\"/Ñ6ÑØC\\*NòH*Mzì¼5y$Å3¥,å¹'©¼L\rLÝ:v m2=:½1qBª!M¶gêgæfvË¬e²þÅn·/Ék³¬Y-\n¶B¦èTZ(×*²geWf¿ÍÊ9«+ÍíÌ³ÊÛ7ïÿíÂá¶¥KW-Xæ½¬j9²<qyÛ\nã+V¬<¸¶*mÕO«íW®~½&zMk^ÁÊÁµkëU\nå}ëÜ×í]OX/Yßµaú>®ÛØ(ÜxåoÊ¿Ü´©«Ä¹dÏfÒféæÞ-[ªæn\rÙÚ´\rßV´íõöEÛ/Í(Û»¶C¹£¿<¸¼e§ÉÎÍ;?T¤TôTúT6îÒÝµa×ønÑî{¼ö4ìÕÛ[¼÷ý>É¾ÛUUMÕfÕeûIû³÷?®ªéøûm]­NmqíÇÒý#¶×¹ÔÕÒ=TRÖ+ëGÇ¾þïw-\r6\rUÆâ#pDyäé÷\tß÷\r:Úv{¬áÓvg/jBòFSû[b[ºOÌ>ÑÖêÞzüGÛ4<YyJóTÉiÚéÓgòÏ}~.ùÜ`Û¢¶{çcÎßjoïºtáÒEÿç;¼;Î\\ò¸tò²ÛåW¸W¯:_mêtê<þÓOÇ»»®¹\\k¹îz½µ{f÷é7ÎÝô½yñÿÖÕ9=Ý½ózo÷Å÷õßÝ~r'ýÎË»Ùw'î­¼O¼_ô@íAÙCÝÕ?[þÜØïÜjÀw óÑÜG÷ÏþõCË\rë8>99â?rýéü§CÏdÏ&þ¢þË®/~øÕë×ÎÑÑ¡ò¿m|¥ýêÀë¯ÛÆÂÆ¾Éx31^ôVûíÁwÜwï£ßOä| (ÿhù±õSÐ§ûÿóüc3-Û cHRMz%ùÿéu0ê`:o_ÅFIDATxÚÏOAÇßÎþ(Mw»KvYcp±Õk(¥ÆH95%&½3I/\tÿG\\1U ñt\r mlS\t%ÛîÌn=l)è;½ù~ò}ßÉ£ªåð?Åü~õrí]»û¼øÏ@½Þ0>yµ¶þüEöàà«ã8N)ß\\÷¼+£Ü¥òÛõÎ#äóyË²fff8[XX¨V«º®[÷î¦¯\r\\är«¹7Ú=LÃáééép8\nÆÆÆ!SSSÝÊÅÇOb7Í°óþC¿,ËL&HÈ²ÜÓÓ£iZ\"ÀB¡ÕÜkÓ4Çieðl6+¢ªªP.  J¥ÒÖÖVtèF­vìñp- ¼ýðÑ³J¥¢iZWW×ë%ÅÍÍÍÝÝÝ÷3Û>q¸bÆxeeåèè!Ôl6mÛE±···P(:CÓ4ZÃ0ù|~nnncc#Àöööàà`.K§ÓKKK,Ëtx;hnF®ë099J¥ !XAèÄ¢(·$¢(÷èóùA°,Ë¸Îí¢(n#F£ñ×]rGUUÓ41Æ$ííí)bÛö¯Sýý}UUëõºi®´¯¯¯R©D£Ñ¶j¯÷§Â÷ùùù`0Ø~6cyyyvvöò¥NUíæyß)P«¿Õ1Æ`Û6!çùÑÑQEæxÞç§özêwR·lÛ9÷4XeúLheE¿ØÂçcE¿À²ìÿX?ªþEÁÄIEND®B`";
        o.save = "PNG\r\n\n\rIHDR_%.-\tpHYsgAMAØëõª cHRMz%ô%Ñm_èl<XçxòIDATxÚbüÿÿ?-@± sdåÌ\\¦þe`r@0Ìÿ!üÝ+ó¸þ00üþûÿË£0³Å`&V1Ýõz_¿ÿaXºp\"CAAÃ¿ÿÀ¬ìÕ«WbbbiMg¾þüùCý+#¯Ü¯ÿ!Áÿÿ±\\ñýÇ_o@|ìø1¸¡ÿÂæÀÇß¿LBû¾þøññ??;#ØLBu1#3ÿåÛ°>y\n1ôß_¸a ð÷ï_0ýáÇo¾ÂIáÿß?Ì¬ ÉÄæbF°Á_ÿ0|úöáÝ»wÿþþºø/Ã/Àô?¨ë¿ýúÊðû×G¨ÁL\"@¾0üûýÈû@,Ñ\tTôéËo/@ÛF­bÄë~ÿøÀðû'Ô`ÿ<@ÈaÒß&?]l¡/Ì µà:ÃÇï¿ÞýÅðèõ_3|ûþbè¿¿:RXþÿÿ\ncFÂêâÏÀ0Þ°ïÃºIVxÓª{!3;rX2ÂXÄn(øül(<L¡a\nï¿°È³ª[ÔL¨¡xb L=òè]XìÃ±AÂ üþþ\tÉÁÖ øøå7<½¢úâbdÿþú5\tÈû/¿à.F\nÉ$ÿÐþóó+V\t#\"¾£<^bw1@aFPd×ülAY#ÇPDá*! FäbSVÅèbI`i¥ÈÀÄ(\n4¡¨àÿ Þ0üûsÿÿ¿ß/>=ù@hÌëÌ@¿ýÿÄ `úÊ@Tqýÿ0}?þþþ ´bX®þÿýhøs >1üc`ZÝüÿ?0Rþÿÿ÷×gpAäÍ&ª2sIEND®B`";
    } else {
        o.bin = "PNG\r\n\n\rIHDRùÇq¦\tpHYsttÞfxÊIDAT(ÏÅ=\n0=cH ¢ ½ä0^À-cü9[Gi¡é`ò}m¥Cßá!¼ÉCðxãñþçcN{°þÊIdÙçù;§ë:­µRªi8Õ¶mAôØ%(pCQeYÖu}Þ5ôØ%ë§Ðß#¥´ÖdgeÙ<Ï$Y'MÓiH²~NduG¬á0$Y§ª*ç¦ôBô¬Ó÷}Bß÷DuÖu½SAÿó[¸Ï@,¢ïT`FIEND®B`";
        o.save = "PNG\r\n\n\rIHDRjß\tpHYsttÞfxÙIDAT8ËÅA\n@=NÁ¡d±Eïã\r¼×h´0hp%#³ñýzeÑ·züóÏÇð½ó<1ÞÏ+Ççã¥cë¸',Öºïû4M7c ç¾¢m[kmÇkfÜý]Ü,á(0#y±<Ïcp¹,Ë ¢Àüx2¡ð}ÿ4A)£¦ir\tE$RJì\\©ëZN@P!lP&Xû\rÊ_~Ï&Ë2BQU2@\reBÑueòWÿfÃ0\\TÜù«m°IEND®B`";
    }
    return o;
}

function setEPubExportOptions() {

    
		var scriptFile = GetActiveScript();
		var scriptFolder = scriptFile.parent;
		var scriptFolderPath = scriptFolder.absoluteURI;
    //set the ISBN to document name without file extention
    // var pattern = "[0-9]+-[0-9]+-[0-9]+-[0-9]+-[0-9]";
    // var theEpubISBN = RegExp(pattern, 'i');

    current_doc.epubExportPreferences.version = EpubVersion.EPUB3;

    //add your publisher metadata between the quotes below
    // current_doc.epubExportPreferences.id = "";
    current_doc.epubExportPreferences.epubTitle = (Folder(current_doc.filePath).displayName);  
    current_doc.epubExportPreferences.epubPublisher = "چشمه";
    current_doc.epubExportPreferences.epubRights= "لیبرا";
    // current_doc.epubExportPreferences.epubCreator = "";
    //current_doc.epubExportPreferences.epubSubject= "";
    //current_doc.epubExportPreferences.epubDate= "";
    //current_doc.epubExportPreferences.epubDescription= "";
    
    current_doc.epubExportPreferences.includeClassesInHTML = true;
    current_doc.epubExportPreferences.ignoreObjectConversionSettings = true;
    current_doc.epubExportPreferences.breakDocument = true;
    current_doc.epubExportPreferences.embedFont = true;
    current_doc.epubExportPreferences.paragraphStyleName = "";
    current_doc.epubExportPreferences.tocStyleName = tocStyle.name; //DefaultTOCStyleName //Epub
    current_doc.epubExportPreferences.stripSoftReturn = true;
    // current_doc.epubExportPreferences.bulletExportOption = BulletListExportOption.UNORDERED_LIST;

    //CascadeStyleSheet
    current_doc.epubExportPreferences.generateCascadeStyleSheet = true;
    current_doc.epubExportPreferences.externalStyleSheets = [File(scriptFolderPath + "/style.css")]; 

    current_doc.epubExportPreferences.exportOrder = ExportOrder.LAYOUT_ORDER;
    current_doc.epubExportPreferences.footnotePlacement = EPubFootnotePlacement.FOOTNOTE_INSIDE_POPUP;

    //Images
    current_doc.epubExportPreferences.imageExportResolution = ImageResolution.PPI_96;
    current_doc.epubExportPreferences.imageConversion = ImageConversion.AUTOMATIC;
    current_doc.epubExportPreferences.imageAlignment = ImageAlignmentType.ALIGN_CENTER;
    current_doc.epubExportPreferences.imagePageBreak = ImagePageBreakType.PAGE_BREAK_BEFORE_AND_AFTER;
    current_doc.epubExportPreferences.customImageSizeOption = ImageSizeOption.SIZE_RELATIVE_TO_TEXT_FLOW;

    current_doc.epubExportPreferences.jpegOptionsQuality = JPEGOptionsQuality.MAXIMUM;
    current_doc.epubExportPreferences.jpegOptionsFormat = JPEGOptionsFormat.PROGRESSIVE_ENCODING;
    current_doc.epubExportPreferences.useSVGAs = UseSVGAsEnum.EMBED_CODE;

    current_doc.epubExportPreferences.gifOptionsPalette = GIFOptionsPalette.ADAPTIVE_PALETTE;
    current_doc.epubExportPreferences.gifOptionsInterlaced = true;

    current_doc.epubExportPreferences.preserveLayoutAppearence = true; 
    current_doc.epubExportPreferences.useExistingImageOnExport = true;
    current_doc.epubExportPreferences.preserveLocalOverride = true;
    current_doc.epubExportPreferences.useImagePageBreak = true;

    //Cover
    current_doc.epubExportPreferences.epubCover = EpubCover.EXTERNAL_IMAGE;
    current_doc.epubExportPreferences.coverImageFile = (File(current_doc.filePath.getFiles("*.jpg")).fsName);
}

function anchorImage(current_doc, textFrame, imageRect) {
    var myAnchoredFrame = CreateAnchor(current_doc, textFrame);
    var imBounds = imageRect.geometricBounds;
    var frBounds = textFrame.geometricBounds
    // Copy image into the anchored frame. Didn't find a better way  
    var imagePath = imageRect.images[0].itemLink.filePath;
    var image = imageRect.images[0];
    myAnchoredFrame.place(File(imagePath));
    myAnchoredFrame.geometricBounds = [imBounds[0] - frBounds[0], imBounds[1] - frBounds[1],
        imBounds[2] - frBounds[0], imBounds[3] - frBounds[1]
    ];
    // Resize image  
    var newImBoundX = myAnchoredFrame.geometricBounds[1] - (imageRect.geometricBounds[1] - image.geometricBounds[1]);
    var newImBoundY = myAnchoredFrame.geometricBounds[0] - (imageRect.geometricBounds[0] - image.geometricBounds[0]);
    var newImBoundX1 = newImBoundX + (image.geometricBounds[3] - image.geometricBounds[1]);
    var newImBoundY1 = newImBoundY + (image.geometricBounds[2] - image.geometricBounds[0]);
    myAnchoredFrame.images[0].geometricBounds = [newImBoundY, newImBoundX, newImBoundY1, newImBoundX1];
    //Set textWrapPreferences of the images  
    myAnchoredFrame.textWrapPreferences.textWrapMode = imageRect.textWrapPreferences.textWrapMode;
    myAnchoredFrame.textWrapPreferences.textWrapOffset = imageRect.textWrapPreferences.textWrapOffset;
    return myAnchoredFrame;
}

function CreateAnchor(current_doc, textFrame) {
    var inPoint = textFrame.insertionPoints[0];
    var anchProps = current_doc.anchoredObjectDefaults.properties;
    var anchCont = anchProps.anchorContent;
    var myAO = inPoint.rectangles.add();
    // Make new object with correct default settings  
    // Make new object right kind of object  
    myAO.contentType = ContentType.graphicType;
    // Recompose parent story so geometricBounds make sense  
    inPoint.parentStory.recompose();
    //save users measurement preferences  
    userHoriz = current_doc.viewPreferences.horizontalMeasurementUnits;
    userVert = current_doc.viewPreferences.verticalMeasurementUnits;
    current_doc.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.points;
    current_doc.viewPreferences.verticalMeasurementUnits = MeasurementUnits.points;
    current_doc.viewPreferences.horizontalMeasurementUnits = userHoriz;
    current_doc.viewPreferences.verticalMeasurementUnits = userVert;
    myAO.applyObjectStyle(anchProps.anchoredObjectStyle);
    if (anchProps.anchorContent == ContentType.textType) {
        try { // might be null  
            myAO.parentStory.appliedParagraphStyle = anchProps.anchoredParagraphStyle;
        } catch (e) {}
    }
    myAO.anchoredObjectSettings.properties = current_doc.anchoredObjectSettings.properties;
    myAO.anchoredObjectSettings.anchoredPosition = AnchorPosition.anchored;
    myAO.anchoredObjectSettings.pinPosition = false;
    return myAO
}

function CountPages() {
    var count = 0;
    // loop through all pages  
    for (i = 0; i < current_doc.pages.length; i++) {
        var page = current_doc.pages.item(i);
        if (page.textFrames.length < 1) continue;
        var textFrame = page.textFrames.item(0);
        if (page.rectangles.length < 1) continue;
        // loop through all rectangles in the page  
        for (j = 0; j < page.rectangles.length; j++) {
            var imageRect = page.rectangles[j];
            if (imageRect.images.length < 1) continue;
            var myAnchoredFrame = anchorImage(current_doc, textFrame, imageRect);
            var pos = [imageRect.geometricBounds[1], imageRect.geometricBounds[0]];
            imageRect.remove();
            j--;
            textFrame.recompose();
            var k = 0;
            // Reposition the anchored image. This is done repeatedly because the first call not moves the  frame to the correct position  
            do {
                myAnchoredFrame.move(pos);
                k++;
            }
            while (k != 5);
            count++;
        }
    }
}

function halfSpace() {
	try {
		var line, findChangeFile, result, findChangeArray, findType, findPreferences, changePreferences, findChangeOptions, comment,
		doc = current_doc;

		var scriptFile = GetActiveScript();
		var scriptFolder = scriptFile.parent;
		var scriptFolderPath = scriptFolder.absoluteURI;
		findChangeFile = new File(scriptFolderPath + "/FindChangeList.txt"); // look into the same folder first
		
		if (!findChangeFile.exists) { // then look into the 'FindChangeSupport' folder located the same folder as the script
			findChangeFile = new File(scriptFolderPath + "/FindChangeSupport/FindChangeList.txt");			
		}
	
		if (!findChangeFile.exists) { // finally look into the default location of the 'samples' scripts
			findChangeFile = new File(app.filePath.absoluteURI + "/Scripts/Scripts Panel/Samples/JavaScript/FindChangeSupport/FindChangeList.txt");
		}		
			
		if (findChangeFile.exists) {
			result = findChangeFile.open("r", undefined, undefined);
			if (result == true) {
				//Loop through the find/change operations.
				do {
					line = findChangeFile.readln();
					//Ignore comment lines and blank lines.
					if ((line.substring(0, 4) == "text") || (line.substring(0, 4) == "grep") || (line.substring(0, 5) == "glyph")) {
						comment = "";
						findChangeArray = line.split("\t");
						//The first field in the line is the findType string.
						findType = findChangeArray[0];
						//The second field in the line is the FindPreferences string.
						findPreferences = findChangeArray[1];
						//The second field in the line is the ChangePreferences string.
						changePreferences = findChangeArray[2];
						//The fourth field is the range--used only by text find/change.
						findChangeOptions = findChangeArray[3];
						//The fifth field is the comment
						if (findChangeArray.length > 4) comment = findChangeArray[4];
						
						switch (findType) {
							case "text":
								FindText(doc, findPreferences, changePreferences, findChangeOptions, comment);
								break;
							case "grep":
								FindGrep(doc, findPreferences, changePreferences, findChangeOptions, comment);
								break;
							case "glyph":
								FindGlyph(doc, findPreferences, changePreferences, findChangeOptions, comment);
								break;
						}
					}
				} while (findChangeFile.eof == false);
				findChangeFile.close();
			}
		}
		else {
			//$.writeln(doc.name + " - Unable to find 'FindChangeList.txt' file.");
			//g.WriteToFile(doc.name + " - Unable to find 'FindChangeList.txt' file.");
		}
	}
	catch(err) {
		//g.WriteToFile(doc.name + " - " + err.message + ", line: " + err.line);
	}
}

function FindText(doc, findPreferences, changePreferences, findChangeOptions, comment) {
	app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
	var str = "app.findTextPreferences.properties = "+ findPreferences + ";";
	str += "app.changeTextPreferences.properties = " + changePreferences + ";";
	str += "app.findChangeTextOptions.properties = " + findChangeOptions + ";";
	app.doScript(str, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT);
	foundItems = doc.changeText();
	if (foundItems.length > 0) {
		//g.WriteToFile("Changed " + foundItems.length + " text item" + ((foundItems.length > 1) ? "s" : "") + ((comment != "") ?  " - " + comment : ""));
	}
	app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
}

function FindGrep(doc, findPreferences, changePreferences, findChangeOptions, comment) {
	app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
	var str = "app.findGrepPreferences.properties = "+ findPreferences + ";";
	str += "app.changeGrepPreferences.properties = " + changePreferences + ";";
	str += "app.findChangeGrepOptions.properties = " + findChangeOptions + ";";
	app.doScript(str, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT);
	var foundItems = doc.changeGrep();
		if (foundItems.length > 0) {
		//g.WriteToFile("Changed " + foundItems.length + " GREP item" + ((foundItems.length > 1) ? "s" : "") + ((comment != "") ?  " - " + comment : ""));
	}
	app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
}

function FindGlyph(doc, findPreferences, changePreferences, findChangeOptions, comment) {
	app.changeGlyphPreferences = app.findGlyphPreferences = NothingEnum.nothing;
	var str = "app.findGlyphPreferences.properties = "+ findPreferences + ";";
	str += "app.changeGlyphPreferences.properties = " + changePreferences + ";";
	str += "app.findChangeGlyphOptions.properties = " + findChangeOptions + ";";
	app.doScript(str, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT);
	var foundItems = doc.changeGlyph();
		if (foundItems.length > 0) {
		//g.WriteToFile("Changed " + foundItems.length + " Glyph item" + ((foundItems.length > 1) ? "s" : "") + ((comment != "") ?  " - " + comment : ""));
	}
	app.changeGlyphPreferences = app.findGlyphPreferences = NothingEnum.NOTHING;
}

function FindFile(filePath) {
	var scriptFile = GetActiveScript();
	var scriptFile = new File(scriptFile);
	var scriptFolder = scriptFile.path;
	var file = new File(scriptFolder + filePath);

	if (file.exists) {
		return file;
	}
	else {
		return null;
	}
}

function GetActiveScript() {
    try {
        return app.activeScript;
    }
    catch(err) {
        return new File(err.fileName);
    }
}

function createTOC() {
    tocStyle = current_doc.tocStyles[0];
    if (tocStyle.tocStyleEntries.length > 0) {  //tocStyle.name == "[Default]" &&
        current_doc.createTOC(tocStyle, true);
        //tocStyle.removeForcedLineBreak = true;
        tocStyle.name = "Epub";
    }
}

