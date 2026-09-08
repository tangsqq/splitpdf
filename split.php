<?php
// Setup environment
ini_set('display_errors', 0);
error_reporting(E_ALL);
ini_set('memory_limit', '1024M');
set_time_limit(300);

// Folder config
$uploadDir = __DIR__ . '/temp_local';
if (!is_dir($uploadDir)) {
    mkdir($uploadDir, 0777, true);
    file_put_contents($uploadDir . '/.htaccess', "Deny from all");
}

// Get Composer
$autoload = __DIR__ . '/vendor/autoload.php';
if (!file_exists($autoload)) {
    header('Content-Type: application/json');
    echo json_encode(['success' => false, 'error' => 'Composer require phpoffice/phpspreadsheet']);
    exit;
}
require $autoload;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;

if (isset($_FILES['excel_file'])) {
    ob_start();
    header('Content-Type: application/json');

    try {
        $file = $_FILES['excel_file'];
        $extension = strtolower(pathinfo($file['name'], PATHINFO_EXTENSION));
        $uniqueId = uniqid();
        $tmpFilePath = $uploadDir . '/' . $uniqueId . '_' . $file['name'];

        if (!move_uploaded_file($file['tmp_name'], $tmpFilePath)) {
            throw new Exception("Upload file fail.");
        }

        if (in_array($extension, ['xlsx', 'xls'])) {
            $spreadsheet = IOFactory::load($tmpFilePath);
            foreach ($spreadsheet->getAllSheets() as $sheet) {
                $setup = $sheet->getPageSetup();
                $setup->setOrientation(PageSetup::ORIENTATION_PORTRAIT);
                $setup->setFitToWidth(1);
                $setup->setFitToHeight(0);
                $setup->setFitToPage(true);
            }
            $writerType = ($extension === 'xls') ? 'Xls' : 'Xlsx';
            $writer = IOFactory::createWriter($spreadsheet, $writerType);
            $writer->save($tmpFilePath);
            unset($spreadsheet, $writer);
        }

        $sofficePath = '"C:\Program Files\LibreOffice\program\soffice.exe"';
        $userProfileDir = $uploadDir . '/profile_' . $uniqueId;
        $userProfile = 'file:///' . str_replace('\\', '/', $userProfileDir);

        $cmd = "$sofficePath \"-env:UserInstallation=$userProfile\" --headless --convert-to pdf --outdir " . escapeshellarg($uploadDir) . " " . escapeshellarg($tmpFilePath) . " 2>&1";
        exec($cmd, $output, $returnVar);

        if ($returnVar === 0) {
            $pdfName = pathinfo($tmpFilePath, PATHINFO_FILENAME) . '.pdf';
            $pdfPath = $uploadDir . '/' . $pdfName;
            if (file_exists($pdfPath)) {
                $base64 = base64_encode(file_get_contents($pdfPath));
                @unlink($pdfPath);
                ob_end_clean();
                echo json_encode(['success' => true, 'pdf_base64' => $base64, 'filename' => $file['name']]);
            } else {
                throw new Exception("PDF convert successfully but can't find file.");
            }
        } else {
            throw new Exception("LibreOffice error: " . implode("\n", $output));
        }
    } catch (Throwable $e) {
        if (ob_get_length()) {
            ob_end_clean();
        }
        echo json_encode(['success' => false, 'error' => $e->getMessage()]);
    }
    if (isset($tmpFilePath) && file_exists($tmpFilePath)) {
        @unlink($tmpFilePath);
    }
    exit;
}
?>
<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <link rel="icon"
        href="data:image/svg+xml,<svg xmlns=%22http://www.w3.org/2000/svg%22 viewBox=%220 0 100 100%22><text y=%22.9em%22 font-size=%2290%22>📑</text></svg>">
    <title>PDF Reorder, Rotate & Split</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.4.120/pdf.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
    <script src="https://unpkg.com/pdf-lib@1.17.1/dist/pdf-lib.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>

    <style>
        :root {
            --primary: #1e293b;
            --primary-hover: #334155;
            --h2-start: #1e293b;
            --h2-end: #334155;
            --split: #f87171;
            --bg: #f8fafc;
            --card-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1), 0 2px 4px -2px rgb(0 0 0 / 0.1);
        }

        body {
            font-family: 'Inter', 'Segoe UI', system-ui, -apple-system, sans-serif;
            background-color: var(--bg);
            background-image: radial-gradient(#e2e8f0 1px, transparent 1px);
            background-size: 20px 20px;
            margin: 0;
            padding: 40px 20px;
            color: #1e293b;
            min-height: 100vh;
        }

        /* Pink Theme */
        body[data-theme="pink"] {
            --primary: #ec4899;
            --primary-hover: #f1d6e2;
            --h2-start: #ec4899;
            --h2-end: #f1d6e2;
            --btn-shadow: rgba(236, 72, 153, 0.2);
        }

        body[data-theme="pink"] .btn-main {
            background: var(--primary) !important;
        }

        body[data-theme="pink"] .setup-card h2 {
            background: var(--primary);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }

        /* Blue Theme */
        body[data-theme="blue"] {
            --primary: #4baff1;
            --primary-hover: #d6e3f1;
            --h2-start: #2db9f5;
            --h2-end: #d6e6f1;
            --btn-shadow: rgba(72, 157, 236, 0.2);
        }

        body[data-theme="blue"] .btn-main {
            background: var(--primary) !important;
        }

        body[data-theme="blue"] .setup-card h2 {
            background: var(--primary);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }

        .btn-primary {
            background: var(--primary) !important;
            box-shadow: 0 4px 12px var(--btn-shadow) !important;
            transition: all 0.3s ease;
        }

        .btn-primary:hover {
            background: var(--primary-hover) !important;
        }

        .setup-card {
            background: rgba(255, 255, 255, 0.9);
            backdrop-filter: blur(10px);
            padding: 30px;
            border-radius: 24px;
            box-shadow: var(--card-shadow);
            text-align: center;
            max-width: 800px;
            margin: 0 auto 40px;
            border: 1px solid var(--primary);
            position: relative;
            animation: cardEntranceIn 0.5s cubic-bezier(0.16, 1, 0.3, 1);
        }

        @keyframes cardEntranceIn {
            from {
                opacity: 0;
                transform: scale(0.95) translateY(-10px);
            }

            to {
                opacity: 1;
                transform: scale(1) translateY(0);
            }
        }

        body.leaving {
            animation: pageOut 0.26s ease forwards;
        }

        @keyframes pageOut {
            to {
                opacity: 0;
                transform: scale(0.98);
            }
        }

        .theme-switcher {
            display: flex;
            gap: 8px;
        }

        .theme-swatch {
            transition: transform 0.25s cubic-bezier(0.34, 1.56, 0.64, 1), box-shadow 0.25s ease;
        }

        .theme-swatch:hover {
            transform: scale(1.25) rotate(8deg);
        }

        .theme-swatch.active {
            transform: scale(1.18);
            box-shadow: 0 0 0 3px rgba(30, 41, 59, 0.2), 0 2px 4px rgba(0, 0, 0, 0.1);
        }

        .help-icon {
            color: #94a3b8;
            cursor: pointer;
            font-size: 18px;
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
            z-index: 10;
        }

        .help-icon:hover {
            color: var(--primary);
            transform: scale(1.2) rotate(15deg);
        }

        @keyframes modalShow {
            from {
                opacity: 0;
                transform: scale(0.9) translateY(-20px);
            }

            to {
                opacity: 1;
                transform: scale(1) translateY(0);
            }
        }

        @keyframes modalHide {
            from {
                opacity: 1;
                transform: scale(1);
            }

            to {
                opacity: 0;
                transform: scale(0.95);
            }
        }

        .modal-overlay {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(15, 23, 42, 0.6);
            backdrop-filter: blur(4px);
            display: none;
            justify-content: center;
            align-items: center;
            z-index: 9999;
            transition: opacity 0.3s ease;
        }

        .modal-animating {
            animation: modalShow 0.4s cubic-bezier(0.34, 1.56, 0.64, 1) forwards;
        }

        .modal-closing {
            animation: modalHide 0.2s ease-in forwards;
        }

        #help-content {
            text-align: left;
            font-size: 14px;
            line-height: 1.6;
            padding: 10px 5px;
        }

        #help-content ul {
            text-align: left;
            margin: 10px 0 0 0;
            padding-left: 0;
            list-style-type: none;
        }

        #help-content li {
            margin-bottom: 10px;
        }

        .setup-card h2 {
            margin: 0 0 30px 0;
            font-weight: 800;
            letter-spacing: -0.025em;
            background: var(--primary);
            -webkit-background-clip: text;
            background-clip: text;
            -webkit-text-fill-color: transparent;
            transition: all 0.3s ease;
        }

        .btn {
            padding: 10px 24px;
            border-radius: 12px;
            border: none;
            cursor: pointer;
            font-weight: 600;
            font-size: 14px;
            transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
            display: inline-flex;
            align-items: center;
            gap: 8px;
            height: 42px;
            box-sizing: border-box;
        }

        .btn-main {
            background: var(--primary) !important;
            color: white;
            border-radius: 20px;
            box-shadow: 0 4px 12px var(--btn-shadow);
        }

        .btn-main:hover {
            background: var(--primary-hover, #334155);
            transform: translateY(-1px);
        }

        .btn-clear {
            background: #fff;
            color: var(--primary, #64748b);
            border: 1px solid var(--primary, #e2e8f0);
            border-radius: 20px;
        }

        .btn-clear:hover {
            background: var(--bg);
            opacity: 0.8;
        }

        #file-selector {
            display: none;
        }

        .file-upload-label {
            padding: 10px 24px;
            border-radius: 20px;
            background: #f1f5f9;
            color: var(--primary, #475569);
            border: 1px solid var(--primary, #e2e8f0);
            cursor: pointer;
            font-weight: 600;
            font-size: 14px;
            transition: all 0.2s;
            display: inline-flex;
            align-items: center;
            gap: 8px;
            height: 42px;
            box-sizing: border-box;
        }

        .workspace-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(230px, 1fr));
            gap: 30px;
            padding: 30px;
            background: rgba(255, 255, 255, 0.5);
            border-radius: 25px;
            border: 2px dashed #cbd5e1;
            min-height: 400px;
            max-width: 1200px;
            margin: 0 auto;
            position: relative;
            counter-reset: page-counter;
            /* Use CSS counter for numbering */
        }

        .drop-hint {
            grid-column: 1 / -1;
            text-align: center;
            color: var(--primary);
            padding-top: 150px;
            pointer-events: none;
        }

        .drop-hint i {
            font-size: 50px;
            margin-bottom: 15px;
            display: block;
            color: var(--primary);
            animation: dropHintFloat 2.2s ease-in-out infinite;
        }

        @keyframes dropHintFloat {

            0%,
            100% {
                transform: translateY(0);
            }

            50% {
                transform: translateY(-10px);
            }
        }

        .workspace-grid.drag-active {
            border-color: var(--primary);
            background: rgba(255, 255, 255, 0.75);
            animation: workspacePulse 0.9s ease-in-out infinite;
        }

        @keyframes workspacePulse {

            0%,
            100% {
                box-shadow: inset 0 0 0 0 rgba(37, 99, 235, 0.15);
            }

            50% {
                box-shadow: inset 0 0 0 8px rgba(37, 99, 235, 0.06);
            }
        }

        /* Show segment header only when split is active on previous card */
        .segment-header {
            grid-column: 1 / -1;
            display: none;
            /* Hide by default */
            align-items: center;
            gap: 12px;
            background: white;
            padding: 15px 25px;
            border-radius: 18px;
            margin-top: 15px;
            border: 1px solid #e2e8f0;
        }

        .segment-header.force-show {
            display: flex;
        }

        /* Trigger header display */
        .page-card.split-active+.segment-header {
            display: flex;
        }

        .rename-input {
            border: 1px solid #cbd5e1;
            border-radius: 10px;
            padding: 8px 15px;
            font-size: 14px;
            flex-grow: 1;
            max-width: 400px;
            outline: none;
        }

        .page-card {
            background: white;
            border: 1px solid #e2e8f0;
            border-radius: 20px;
            padding: 12px;
            cursor: grab;
            position: relative;
            transition: border-color 0.3s, transform 0.3s;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.02);
            user-select: none;
            animation: pageCardIn 0.35s cubic-bezier(0.34, 1.56, 0.64, 1) backwards;
        }

        @keyframes pageCardIn {
            from {
                opacity: 0;
                transform: scale(0.88) translateY(8px);
            }

            to {
                opacity: 1;
                transform: scale(1) translateY(0);
            }
        }

        .page-card:hover {
            border-color: var(--primary);
            transform: translateY(-5px);
        }

        .page-card.selected {
            border: 2px solid var(--primary);
            background: #eff6ff;
            box-shadow: 0 0 15px rgba(37, 99, 235, 0.2);
        }

        canvas {
            width: 100%;
            height: auto;
            border-radius: 20px;
            display: block;
            transition: transform 0.4s cubic-bezier(0.34, 1.56, 0.64, 1);
            background: #f8fafc;
        }

        .badge {
            position: absolute;
            top: -12px;
            left: 12px;
            background: var(--primary);
            color: white;
            padding: 4px 10px;
            border-radius: 20px;
            font-size: 11px;
            font-weight: 700;
            z-index: 5;
            transition: background 0.3s ease;
        }

        /* Auto-generate page number badge */
        .badge::after {
            counter-increment: page-counter;
            content: "#" counter(page-counter);
        }

        .rotate-btn {
            position: absolute;
            bottom: 35px;
            right: 18px;
            border: 1px solid #e2e8f0;
            border-radius: 50%;
            width: 32px;
            height: 32px;
            cursor: pointer;
            display: flex;
            align-items: center;
            justify-content: center;
            z-index: 10;
            background: white;
            color: var(--primary);
        }

        .rotate-btn:hover {
            border-color: var(--primary, #1e293b);
            transform: rotate(25deg);
        }

        .rotate-btn {
            transition: transform 0.25s ease, border-color 0.25s ease;
        }

        .rotate-btn.spin-once {
            animation: rotateSpin 0.4s ease;
        }

        @keyframes rotateSpin {
            from {
                transform: rotate(0deg);
            }

            to {
                transform: rotate(180deg);
            }
        }

        .delete-btn {
            transition: transform 0.2s cubic-bezier(0.34, 1.56, 0.64, 1);
        }

        .delete-btn:hover {
            transform: scale(1.15) rotate(90deg);
        }

        .delete-btn {
            position: absolute;
            top: -12px;
            right: 12px;
            background: var(--split);
            color: white;
            border: none;
            border-radius: 50%;
            width: 24px;
            height: 24px;
            cursor: pointer;
            display: none;
            align-items: center;
            justify-content: center;
            z-index: 15;
        }

        .page-card:hover .delete-btn {
            display: flex;
        }

        .page-card.split-active {
            border-right: 4px dashed var(--split);
            margin-right: 5px;
        }

        .page-card.split-active::after {
            content: '✂️';
            position: absolute;
            right: -16px;
            top: 50%;
            transform: translateY(-50%);
            background: white;
            border-radius: 50%;
            padding: 2px;
        }

        .modal-overlay {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(15, 23, 42, 0.6);
            backdrop-filter: blur(4px);
            display: none;
            justify-content: center;
            align-items: center;
            z-index: 9999;
        }

        #previewModal {
            background: rgba(0, 0, 0, 0.95);
            overflow: hidden;
        }

        #previewImage {
            max-width: 90%;
            max-height: 90vh;
            box-shadow: 0 0 40px rgba(0, 0, 0, 0.8);
            cursor: grab;
            user-select: none;
            transition: transform 0.1s ease-out;
        }

        #previewImage:active {
            cursor: grabbing;
        }

        .nav-arrow {
            position: absolute;
            top: 50%;
            transform: translateY(-50%);
            background: rgba(255, 255, 255, 0.1);
            color: white;
            border: none;
            width: 60px;
            height: 60px;
            border-radius: 50%;
            cursor: pointer;
            font-size: 24px;
            display: flex;
            align-items: center;
            justify-content: center;
            transition: background 0.2s;
            z-index: 10001;
        }

        .nav-arrow:hover {
            background: rgba(255, 255, 255, 0.3);
        }

        #prevArrow {
            left: 30px;
        }

        #nextArrow {
            right: 30px;
        }

        .home-btn {
            position: fixed;
            top: 20px;
            right: 20px;
            color: var(--primary, #1e293b);
            text-decoration: none;
            transition: all 0.3s ease;
            z-index: 100;
        }

        .home-btn:hover {
            transform: scale(1.1);
            color: var(--primary-hover);
        }

        .home-btn i {
            font-size: 30px;
            display: inline-block;
            animation: homeFloat 3s ease-in-out infinite;
        }

        .home-btn:hover i {
            animation: none;
            transform: rotate(-10deg) scale(1.12);
        }

        @keyframes homeFloat {

            0%,
            100% {
                transform: translateY(0);
            }

            50% {
                transform: translateY(-3px);
            }
        }

        .top-right-controls {
            position: absolute;
            top: 15px;
            right: 15px;
            display: flex;
            align-items: center;
            gap: 15px;
            z-index: 10;
        }
    </style>
</head>

<body>
    <div class="setup-card">
        <div class="top-right-controls">
            <div class="theme-switcher">
                <button onclick="setTheme('default')" title="Default Theme"
                    style="background:#1e293b; width:16px; height:16px; border-radius:50%; border:2px solid #fff; cursor:pointer; box-shadow: 0 2px 4px rgba(0,0,0,0.1);"></button>
                <button onclick="setTheme('pink')" title="Pink Theme"
                    style="background:#ec4899; width:16px; height:16px; border-radius:50%; border:2px solid #fff; cursor:pointer; box-shadow: 0 2px 4px rgba(0,0,0,0.1);"></button>
                <button onclick="setTheme('blue')" title="Blue Theme"
                    style="background:#4baff1; width:16px; height:16px; border-radius:50%; border:2px solid #fff; cursor:pointer; box-shadow: 0 2px 4px rgba(0,0,0,0.1);"></button>
            </div>
            <i class="fa-regular fa-circle-question help-icon" title="How to use" onclick="showHelp()"></i>
        </div>
        <h2>PDF Reorder, Rotate & Split</h2>
        <div style="display: flex; align-items: center; justify-content: center; gap: 12px; flex-wrap: wrap;">
            <label for="file-selector" class="file-upload-label"
                style="border-color: var(--primary); color: var(--primary);">
                <svg width="20" height="20" fill="none" stroke="currentColor" viewBox="0 0 24 24"
                    xmlns="http://www.w3.org/2000/svg">
                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 4v16m8-8H4"></path>
                </svg>
                Choose Files
            </label>
            <input type="file" id="file-selector"
                accept="application/pdf, .xlsx, .xls, .doc, .docx, .ppt, .pptx, .jpg, .jpeg, .png" multiple>
            <button class="btn btn-main" onclick="openPdfEditor()">Edit PDF</button>
            <button class="btn btn-main" onclick="exportPDF()">Download All</button>
            <button class="btn btn-clear" onclick="location.reload()">Clear All</button>
        </div>
    </div>

    <div id="workspace" class="workspace-grid">
        <div class="drop-hint" id="drop-hint"><i class="fa-solid fa-file-arrow-up"></i>Drag and Drop files here</div>
    </div>

    <div id="customAlert" class="modal-overlay">
        <div
            style="background: white; padding: 32px; border-radius: 25px; text-align: center; max-width: 360px; width: 90%;">
            <h3 id="alertTitle">Status</h3>
            <p id="alertMessage"></p>
            <button class="btn btn-main" id="alertBtn" onclick="closeAlert()">OK</button>
        </div>
    </div>

    <div id="previewModal" class="modal-overlay" onclick="closePreview()">
        <button class="nav-arrow" id="prevArrow" onclick="navigatePreview(-1, event)"><i
                class="fa fa-chevron-left"></i></button>
        <img id="previewImage" src="" onclick="event.stopPropagation()">
        <button class="nav-arrow" id="nextArrow" onclick="navigatePreview(1, event)"><i
                class="fa fa-chevron-right"></i></button>
    </div>

    <a href="index.html" class="home-btn" title="Back to Home"><i class="fa fa-home"></i></a>


    <style>
        /* Adobe-style PDF editor */
        #pdfEditorModal {
            position: fixed;
            inset: 0;
            background: rgba(15, 23, 42, .94);
            z-index: 20000;
            display: none;
            flex-direction: column;
            color: #fff;
        }

        #pdfEditorModal.open {
            display: flex;
        }

        .pdf-editor-topbar {
            height: 58px;
            flex: 0 0 58px;
            display: flex;
            align-items: center;
            gap: 8px;
            padding: 0 14px;
            background: #111827;
            border-bottom: 1px solid #374151;
        }

        .pdf-editor-title {
            font-weight: 700;
            margin-right: auto;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
            max-width: 280px;
        }

        .pdf-editor-tool {
            border: 1px solid #4b5563;
            background: #1f2937;
            color: #e5e7eb;
            border-radius: 8px;
            height: 36px;
            min-width: 38px;
            padding: 0 11px;
            cursor: pointer;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            gap: 6px;
        }

        .pdf-editor-tool:hover,
        .pdf-editor-tool.active {
            background: #374151;
            border-color: #94a3b8;
        }

        .pdf-editor-tool.close {
            background: #7f1d1d;
            border-color: #991b1b;
        }

        .pdf-editor-stage {
            flex: 1;
            min-height: 0;
            display: flex;
            align-items: center;
            justify-content: center;
            overflow: auto;
            padding: 24px;
        }

        #pdfEditorCanvas {
            background: white;
            box-shadow: 0 12px 45px rgba(0, 0, 0, .45);
            cursor: crosshair;
            max-width: none;
        }

        .pdf-editor-status {
            height: 42px;
            flex: 0 0 42px;
            background: #111827;
            border-top: 1px solid #374151;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 14px;
            font-size: 13px;
        }

        .pdf-editor-status button {
            border: 0;
            background: transparent;
            color: #e5e7eb;
            cursor: pointer;
        }

        .pdf-editor-color {
            width: 30px;
            height: 30px;
            border: 0;
            padding: 0;
            background: transparent;
            cursor: pointer;
        }

        .pdf-editor-size {
            width: 78px;
            accent-color: #60a5fa;
        }

        .pdf-editor-help {
            color: #9ca3af;
            font-size: 12px;
        }

        @media (max-width: 900px) {
            .pdf-editor-tool span {
                display: none;
            }

            .pdf-editor-title {
                max-width: 120px;
            }

            .pdf-editor-stage {
                padding: 10px;
            }
        }
    </style>

    <!-- PDF Editor modal -->
    <div id="pdfEditorModal">
        <div class="pdf-editor-topbar">
            <button class="pdf-editor-tool" onclick="editorPreviousPage()" title="Previous page">
                <i class="fa fa-chevron-left"></i>
            </button>
            <button class="pdf-editor-tool" onclick="editorNextPage()" title="Next page">
                <i class="fa fa-chevron-right"></i>
            </button>
            <div class="pdf-editor-title" id="pdfEditorTitle">Edit PDF</div>

            <button class="pdf-editor-tool active" data-editor-tool="select" onclick="setEditorTool('select')"
                title="Select">
                <i class="fa fa-arrow-pointer"></i><span>Select</span>
            </button>
            <button class="pdf-editor-tool" data-editor-tool="text" onclick="setEditorTool('text')" title="Add text">
                <i class="fa fa-font"></i><span>Text</span>
            </button>
            <button class="pdf-editor-tool" data-editor-tool="highlight" onclick="setEditorTool('highlight')"
                title="Highlight">
                <i class="fa fa-highlighter"></i><span>Highlight</span>
            </button>
            <button class="pdf-editor-tool" data-editor-tool="draw" onclick="setEditorTool('draw')" title="Draw">
                <i class="fa fa-pen"></i><span>Draw</span>
            </button>
            <button class="pdf-editor-tool" data-editor-tool="whiteout" onclick="setEditorTool('whiteout')"
                title="Whiteout">
                <i class="fa fa-eraser"></i><span>Whiteout</span>
            </button>
            <input id="pdfEditorColor" class="pdf-editor-color" type="color" value="#ef4444" title="Color">
            <input id="pdfEditorSize" class="pdf-editor-size" type="range" min="1" max="20" value="3"
                title="Brush/text size">

            <button class="pdf-editor-tool" onclick="undoEditorEdit()" title="Undo">
                <i class="fa fa-rotate-left"></i>
            </button>
            <button class="pdf-editor-tool" onclick="clearEditorPage()" title="Clear edits on this page">
                <i class="fa fa-trash"></i>
            </button>
            <button class="pdf-editor-tool" onclick="editorZoom(-1)" title="Zoom out">
                <i class="fa fa-minus"></i>
            </button>
            <button class="pdf-editor-tool" onclick="editorZoom(1)" title="Zoom in">
                <i class="fa fa-plus"></i>
            </button>
            <button class="pdf-editor-tool" onclick="saveEditorChanges()" title="Save edits">
                <i class="fa fa-check"></i><span>Save</span>
            </button>
            <button class="pdf-editor-tool close" onclick="closePdfEditor()" title="Close">
                <i class="fa fa-times"></i><span>Close</span>
            </button>
        </div>

        <div class="pdf-editor-stage" id="pdfEditorStage">
            <canvas id="pdfEditorCanvas"></canvas>
        </div>

        <div class="pdf-editor-status">
            <span id="pdfEditorPageInfo">Page 1 / 1</span>
            <span class="pdf-editor-help">Text: click page • Highlight/Whiteout: drag • Draw: drag</span>
        </div>
    </div>

    <script>
        const {
            PDFDocument,
            degrees
        } = PDFLib;
        pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.4.120/pdf.worker.min.js';

        let sourcePdfs = new Map();
        let currentPreviewIdx = -1;
        let state = {
            pageOrder: [],
            splits: new Set(),
            segmentNames: {},
            selectedIndices: new Set()
        };

        // PDF editing state. Edits are stored separately from page thumbnails so
        // they can be applied to the final PDF without changing the source file.
        const pdfEdits = new Map();
        let editorPageIndex = 0;
        let editorTool = 'select';
        let editorZoomLevel = 1;
        let editorDrawing = false;
        let editorStart = null;
        let editorCurrentPath = [];
        let editorBaseImage = null;
        let zoomLevel = 1,
            isDragging = false,
            startX, startY, translateX = 0,
            translateY = 0;

        // Theme switching logic
        function setTheme(theme) {

            if (theme === 'pink') {

                document.body.setAttribute('data-theme', 'pink');
                localStorage.setItem('selected-theme', 'pink');

            } else if (theme === 'blue') {

                document.body.setAttribute('data-theme', 'blue');
                localStorage.setItem('selected-theme', 'blue');

            } else {

                document.body.removeAttribute('data-theme');
                localStorage.setItem('selected-theme', 'default');
            }
        }

        // Apply saved theme on page load
        document.addEventListener('DOMContentLoaded', () => {
            const savedTheme = localStorage.getItem('selected-theme') || 'default';
            setTheme(savedTheme);
        });

        function showHelp() {
            const helpText = `
            <div id="help-content">
                <strong>Main Features:</strong>
                <ul>
                    <li><strong>Upload:</strong> Supports PDF, Office (Word, Excel, PPT), and Images. Non-pdf will automatically convert to PDF.</li>
                    <li><strong>Reorder:</strong> Drag and drop pages to change page sequence. Multi-select with Ctrl + Click.</li>
                    <li><strong>Delete:</strong> Delete unwanted pages.</li>
                    <li><strong>Rotate:</strong> Click the rotate icon on any page card.</li>
                    <li><strong>Split:</strong> Right-click to split.</li>
                    <li><strong>Merge:</strong> Merge pages or files from multiple documents into one file.</li>
                    <li><strong>Preview:</strong> Left-click to preview. Scroll to zoom in and out, drag to move around.</li>
                    <li><strong>Rename:</strong> Enter custom names for each split segment in the input fields.</li>
                    <li><strong>Download:</strong> Download individual segments or "Download All" as a ZIP file.</li>
                </ul>
            </div>
        `;

            const alertModal = document.getElementById('customAlert');
            const modalContent = alertModal.querySelector('div');

            document.getElementById('alertTitle').innerText = "How to Use";
            document.getElementById('alertMessage').innerHTML = helpText; // Use innerHTML for the list

            // Adjust alert width for better reading
            modalContent.style.maxWidth = "500px";

            // Trigger animation
            alertModal.style.display = 'flex';
            modalContent.classList.remove('modal-closing');
            modalContent.classList.add('modal-animating');
        }

        function closeAlert() {
            const alertModal = document.getElementById('customAlert');
            const modalContainer = alertModal.querySelector('div');

            modalContainer.classList.remove('modal-animating');
            modalContainer.classList.add('modal-closing');

            setTimeout(() => {
                alertModal.style.display = 'none';
                modalContainer.classList.remove('modal-closing');

                document.getElementById('alertTitle').innerText = "Status";
                modalContainer.style.maxWidth = "360px";
            }, 200);
        }

        function generateFileId() {
            return Math.random().toString(36).substring(2, 15) + Date.now().toString(36);
        }

        function cloneBuffer(buffer) {
            const dst = new ArrayBuffer(buffer.byteLength);
            new Uint8Array(dst).set(new Uint8Array(buffer));
            return dst;
        }

        const workspace = document.getElementById('workspace');

        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            document.addEventListener(eventName, e => {
                e.preventDefault();
                e.stopPropagation();
            }, false);
        });

        workspace.addEventListener('drop', (e) => {
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                // If dropping directly on workspace (hint area), append to end
                if (e.target === workspace || e.target.id === 'drop-hint') {
                    handleFiles(Array.from(files));
                }
            }
        });

        document.getElementById('file-selector').addEventListener('change', (e) => {
            handleFiles(Array.from(e.target.files));
        });

        async function handleFiles(files, insertIndex = -1) {
            if (files.length === 0) return;
            showAlert("Processing files...");

            const newPages = [];
            const newFileIds = [];

            for (const file of files) {
                const fileId = generateFileId();
                try {
                    let rawBuffer;
                    const needsConversion = file.name.match(/\.(xlsx|xls|doc|docx|ppt|pptx|jpg|jpeg|png)$/i);
                    if (needsConversion) {
                        const formData = new FormData();
                        formData.append('excel_file', file);
                        const response = await fetch('', {
                            method: 'POST',
                            body: formData
                        });
                        const result = await response.json();
                        if (!result.success) throw new Error(result.error);
                        const binaryStr = atob(result.pdf_base64);
                        const bytes = new Uint8Array(binaryStr.length);
                        for (let i = 0; i < binaryStr.length; i++) bytes[i] = binaryStr.charCodeAt(i);
                        rawBuffer = bytes.buffer;
                    } else if (file.type === "application/pdf") {
                        rawBuffer = await file.arrayBuffer();
                    } else continue;

                    sourcePdfs.set(fileId, {
                        buffer: cloneBuffer(rawBuffer),
                        pdfjsDoc: null
                    });
                    const pdfjsDoc = await pdfjsLib.getDocument({
                        data: cloneBuffer(rawBuffer)
                    }).promise;
                    sourcePdfs.get(fileId).pdfjsDoc = pdfjsDoc;

                    for (let i = 0; i < pdfjsDoc.numPages; i++) {
                        newPages.push({
                            fileId,
                            originalIdx: i,
                            rotation: 0,
                            fileName: file.name
                        });
                    }
                } catch (err) {
                    alert("Error: " + err.message);
                }
            }

            if (insertIndex === -1) {
                state.pageOrder.push(...newPages);
            } else {
                state.pageOrder.splice(insertIndex, 0, ...newPages);
            }

            renderWorkspace();
            closeAlert();
        }

        function renderWorkspace() {
            const fragment = document.createDocumentFragment();
            workspace.innerHTML = '';

            if (state.pageOrder.length === 0) {
                workspace.innerHTML = '<div class="drop-hint" id="drop-hint"><i class="fa-solid fa-file-arrow-up"></i>Drag and Drop files here</div>';
                return;
            }

            const firstHeader = createRenameBar(-1);
            firstHeader.classList.add('force-show');
            fragment.appendChild(firstHeader);

            state.pageOrder.forEach((p, i) => {
                const card = createPageCard(p, i);
                fragment.appendChild(card);
                fragment.appendChild(createRenameBar(i));

                const canvasId = `canvas-${p.fileId}-${p.originalIdx}-${i}`;
                drawThumb(p, canvasId);
            });

            workspace.appendChild(fragment);
        }

        function createRenameBar(idx) {
            const div = document.createElement('div');
            div.className = 'segment-header';
            div.dataset.forIdx = idx;

            const fileNumber = idx + 2;
            const defaultName = `Page_${fileNumber}`;

            div.innerHTML = `
                <span class="segment-label">File Name:</span>
                <input type="text" class="rename-input" value="${state.segmentNames[idx + 1] || ''}" placeholder="${defaultName}" oninput="state.segmentNames[${idx + 1}] = this.value">
                <button class="btn btn-main" style="height:34px; font-size:12px;" onclick="downloadSingleGroupFromDOM(${idx})">Download</button>
            `;
            return div;
        }

        function createPageCard(pageObj, index) {
            const card = document.createElement('div');
            card.className = 'page-card';
            card.draggable = true;
            if (state.selectedIndices.has(index)) card.classList.add('selected');
            if (state.splits.has(index)) card.classList.add('split-active');

            const canvasId = `canvas-${pageObj.fileId}-${pageObj.originalIdx}-${index}`;
            card.innerHTML = `
                <div class="badge"></div>
                <button class="delete-btn" title="Delete"><i class="fa fa-times"></i></button>
                <canvas id="${canvasId}" style="transform: rotate(${pageObj.rotation}deg)"></canvas>
                <button class="rotate-btn" title="Rotate"><i class="fa fa-rotate-right"></i></button>
                <div style="font-size:11px; color:#94a3b8; margin-top:8px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">${pageObj.fileName}</div>
            `;

            card.onclick = (e) => {
                if (e.ctrlKey || e.metaKey) {
                    if (state.selectedIndices.has(index)) {
                        state.selectedIndices.delete(index);
                        card.classList.remove('selected');
                    } else {
                        state.selectedIndices.add(index);
                        card.classList.add('selected');
                    }
                } else {
                    showPreview(index);
                }
            };

            card.oncontextmenu = (e) => {
                e.preventDefault();
                if (index === state.pageOrder.length - 1) return;

                if (card.classList.contains('split-active')) {
                    card.classList.remove('split-active');
                    state.splits.delete(index);
                } else {
                    card.classList.add('split-active');
                    state.splits.add(index);
                }
            };

            card.querySelector('.rotate-btn').onclick = (e) => {
                e.stopPropagation();
                pageObj.rotation = (pageObj.rotation + 90) % 360;
                card.querySelector('canvas').style.transform = `rotate(${pageObj.rotation}deg)`;
            };

            card.querySelector('.delete-btn').onclick = (e) => {
                e.stopPropagation();
                state.pageOrder.splice(index, 1);
                state.selectedIndices.clear(); // Clear to avoid index mismatch
                renderWorkspace();
            };

            card.ondragstart = (e) => {
                if (!state.selectedIndices.has(index)) {
                    state.selectedIndices.clear();
                    state.selectedIndices.add(index);
                }
                e.dataTransfer.setData('text/plain', 'multi');
            };

            card.ondragover = (e) => {
                e.preventDefault();
                const rect = card.getBoundingClientRect();
                const relX = e.clientX - rect.left;
                card.style.borderLeft = (relX < rect.width / 2) ? "4px solid var(--primary)" : "";
                card.style.borderRight = (relX >= rect.width / 2) ? "4px solid var(--primary)" : "";
            };

            card.ondragleave = () => {
                card.style.borderLeft = "";
                card.style.borderRight = "";
            };

            card.ondrop = (e) => {
                e.preventDefault();
                card.style.borderLeft = "";
                card.style.borderRight = "";

                const rect = card.getBoundingClientRect();
                const relX = e.clientX - rect.left;
                const dropInsertIndex = (relX < rect.width / 2) ? index : index + 1;

                if (e.dataTransfer.files.length > 0) {
                    // Handle external files drop between pages
                    handleFiles(Array.from(e.dataTransfer.files), dropInsertIndex);
                } else {
                    // Handle internal reordering
                    const sortedSelected = Array.from(state.selectedIndices).sort((a, b) => a - b);
                    const itemsToMove = sortedSelected.map(i => state.pageOrder[i]);

                    // Remove items
                    for (let i = sortedSelected.length - 1; i >= 0; i--) {
                        state.pageOrder.splice(sortedSelected[i], 1);
                    }

                    // Calculate new insert position
                    let finalInsert = dropInsertIndex;
                    const shift = sortedSelected.filter(i => i < dropInsertIndex).length;
                    finalInsert -= shift;

                    state.pageOrder.splice(finalInsert, 0, ...itemsToMove);
                    state.selectedIndices.clear();
                    renderWorkspace();
                }
            };

            return card;
        }

        async function drawThumb(pageObj, canvasId) {
            try {
                const source = sourcePdfs.get(pageObj.fileId);
                const page = await source.pdfjsDoc.getPage(pageObj.originalIdx + 1);
                const viewport = page.getViewport({
                    scale: 1.5
                });
                const canvas = document.getElementById(canvasId);
                if (!canvas) return;
                const ctx = canvas.getContext('2d');
                canvas.height = viewport.height;
                canvas.width = viewport.width;
                await page.render({
                    canvasContext: ctx,
                    viewport
                }).promise;
            } catch (e) {
                console.error(e);
            }
        }

        function showPreview(index) {
            currentPreviewIdx = index;
            zoomLevel = 1;
            translateX = 0;
            translateY = 0;
            const cards = workspace.querySelectorAll('.page-card');
            const sourceCanvas = cards[index].querySelector('canvas');

            const img = document.getElementById('previewImage');
            img.src = sourceCanvas.toDataURL('image/png');
            updateImageTransform();

            document.getElementById('previewModal').style.display = 'flex';
            document.getElementById('prevArrow').style.visibility = index > 0 ? 'visible' : 'hidden';
            document.getElementById('nextArrow').style.visibility = index < state.pageOrder.length - 1 ? 'visible' : 'hidden';
        }

        function updateImageTransform() {
            const img = document.getElementById('previewImage');
            const pageObj = state.pageOrder[currentPreviewIdx];
            const rotation = pageObj ? `rotate(${pageObj.rotation}deg)` : 'rotate(0deg)';
            img.style.transform = `translate(${translateX}px, ${translateY}px) ${rotation} scale(${zoomLevel})`;
        }

        function navigatePreview(direction, event) {
            event.stopPropagation();
            const newIdx = currentPreviewIdx + direction;
            if (newIdx >= 0 && newIdx < state.pageOrder.length) showPreview(newIdx);
        }

        function closePreview() {
            document.getElementById('previewModal').style.display = 'none';
        }

        // Export Logic 
        function splitIntoGroups() {
            const cards = Array.from(workspace.querySelectorAll('.page-card'));
            const groups = [];
            let currentGroup = [];

            cards.forEach((card, i) => {
                currentGroup.push(state.pageOrder[i]);
                if (card.classList.contains('split-active')) {
                    groups.push(currentGroup);
                    currentGroup = [];
                }
            });
            groups.push(currentGroup);
            return groups;
        }

        async function downloadSingleGroupFromDOM(headerIdx) {
            showAlert("Preparing download...");
            const groups = splitIntoGroups();
            const headers = Array.from(workspace.querySelectorAll('.segment-header')).filter(h =>
                h.classList.contains('force-show') || h.previousElementSibling.classList.contains('split-active')
            );
            const groupIdx = headers.findIndex(h => h.dataset.forIdx == headerIdx || (headerIdx === -1 && h.classList.contains('force-show')));

            const bytes = await generatePdfBlob(groups[groupIdx]);
            const nameInput = headers[groupIdx].querySelector('.rename-input');
            downloadBlob(bytes, nameInput.value || `Document_Page_${groupIdx + 1}`);
            closeAlert();
        }

        async function exportPDF() {
            if (state.pageOrder.length === 0) return showAlert("No pages to export.");
            showAlert("Generating ZIP file...");
            const zip = new JSZip();
            const groups = splitIntoGroups();
            const headers = Array.from(workspace.querySelectorAll('.segment-header')).filter(h =>
                h.classList.contains('force-show') || h.previousElementSibling.classList.contains('split-active')
            );

            for (let i = 0; i < groups.length; i++) {
                const bytes = await generatePdfBlob(groups[i]);
                const name = headers[i].querySelector('.rename-input').value || `Document_Page_${i + 1}`;
                zip.file(name.toLowerCase().endsWith('.pdf') ? name : name + '.pdf', bytes);
            }
            const content = await zip.generateAsync({
                type: "blob"
            });
            const link = document.createElement('a');
            link.href = URL.createObjectURL(content);
            link.download = "converted_documents.zip";
            link.click();
            closeAlert();
        }

        async function generatePdfBlob(group) {
            const newDoc = await PDFDocument.create();
            const cache = new Map();
            for (const pageObj of group) {
                if (!cache.has(pageObj.fileId)) cache.set(pageObj.fileId, await PDFDocument.load(cloneBuffer(sourcePdfs.get(pageObj.fileId).buffer)));
                const srcDoc = cache.get(pageObj.fileId);
                const [copiedPage] = await newDoc.copyPages(srcDoc, [pageObj.originalIdx]);

                // Apply Adobe-style overlay edits (text, highlight, drawing and whiteout).
                await applyPdfEdits(copiedPage, pageObj);

                if (pageObj.rotation !== 0) copiedPage.setRotation(degrees(pageObj.rotation));
                newDoc.addPage(copiedPage);
            }
            return await newDoc.save();
        }

        function downloadBlob(bytes, filename) {
            const blob = new Blob([bytes], {
                type: 'application/pdf'
            });
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = filename.toLowerCase().endsWith('.pdf') ? filename : filename + '.pdf';
            a.click();
        }

        function showAlert(msg) {
            document.getElementById('alertMessage').innerText = msg;
            document.getElementById('customAlert').style.display = 'flex';
        }

        function closeAlert() {
            document.getElementById('customAlert').style.display = 'none';
        }

        // Preview zoom and pan
        document.getElementById('previewModal').addEventListener('wheel', (e) => {
            e.preventDefault();
            zoomLevel = Math.min(Math.max(0.5, zoomLevel + (e.deltaY > 0 ? -0.1 : 0.1)), 5);
            updateImageTransform();
        }, {
            passive: false
        });

        document.getElementById('previewImage').addEventListener('mousedown', (e) => {
            isDragging = true;
            startX = e.clientX - translateX;
            startY = e.clientY - translateY;
        });
        window.addEventListener('mousemove', (e) => {
            if (!isDragging) return;
            translateX = e.clientX - startX;
            translateY = e.clientY - startY;
            updateImageTransform();
        });

        /* ==================== PDF EDITOR ==================== */

        function getEditorKey(pageObj) {
            return `${pageObj.fileId}:${pageObj.originalIdx}`;
        }

        function getEditorEdits(pageObj) {
            const key = getEditorKey(pageObj);
            if (!pdfEdits.has(key)) pdfEdits.set(key, []);
            return pdfEdits.get(key);
        }

        function openPdfEditor() {
            if (!state.pageOrder.length) {
                showAlert("Please add a PDF first.");
                return;
            }

            // Edit the currently selected page, otherwise start at page 1.
            const selected = Array.from(state.selectedIndices).sort((a, b) => a - b);
            editorPageIndex = selected.length ? selected[0] : 0;
            editorZoomLevel = 1;
            editorTool = 'select';

            document.getElementById('pdfEditorModal').classList.add('open');
            setEditorTool('select');
            renderEditorPage();
        }

        function closePdfEditor() {
            document.getElementById('pdfEditorModal').classList.remove('open');
            editorDrawing = false;
            editorStart = null;
            editorCurrentPath = [];
        }

        function saveEditorChanges() {
            closePdfEditor();
            renderWorkspace();
            showAlert("PDF edits saved. Use Download All to export them.");
            setTimeout(closeAlert, 1200);
        }

        function setEditorTool(tool) {
            editorTool = tool;
            document.querySelectorAll('[data-editor-tool]').forEach(btn => {
                btn.classList.toggle('active', btn.dataset.editorTool === tool);
            });

            const canvas = document.getElementById('pdfEditorCanvas');
            canvas.style.cursor = tool === 'select' ? 'default' : 'crosshair';
        }

        function editorPreviousPage() {
            if (editorPageIndex > 0) {
                editorPageIndex--;
                renderEditorPage();
            }
        }

        function editorNextPage() {
            if (editorPageIndex < state.pageOrder.length - 1) {
                editorPageIndex++;
                renderEditorPage();
            }
        }

        function editorZoom(direction) {
            editorZoomLevel = Math.min(3, Math.max(.5, editorZoomLevel + direction * .15));
            renderEditorPage();
        }

        async function renderEditorPage() {
            const pageObj = state.pageOrder[editorPageIndex];
            if (!pageObj) return;

            const source = sourcePdfs.get(pageObj.fileId);
            if (!source || !source.pdfjsDoc) return;

            const page = await source.pdfjsDoc.getPage(pageObj.originalIdx + 1);
            const baseScale = Math.min(
                1.7,
                Math.max(
                    .8,
                    (window.innerHeight - 150) / page.getViewport({ scale: 1 }).height
                )
            );
            const scale = baseScale * editorZoomLevel;
            const viewport = page.getViewport({ scale, rotation: 0 });

            const canvas = document.getElementById('pdfEditorCanvas');
            canvas.width = Math.ceil(viewport.width);
            canvas.height = Math.ceil(viewport.height);

            const ctx = canvas.getContext('2d');
            ctx.clearRect(0, 0, canvas.width, canvas.height);
            await page.render({ canvasContext: ctx, viewport }).promise;

            editorBaseImage = ctx.getImageData(0, 0, canvas.width, canvas.height);

            drawEditorOverlays(ctx, pageObj, viewport);

            document.getElementById('pdfEditorTitle').textContent =
                `Edit: ${pageObj.fileName || 'PDF'}`;
            document.getElementById('pdfEditorPageInfo').textContent =
                `Page ${editorPageIndex + 1} / ${state.pageOrder.length}`;
        }

        function drawEditorOverlays(ctx, pageObj, viewport) {
            const edits = getEditorEdits(pageObj);

            for (const edit of edits) {
                if (edit.type === 'text') {
                    const fontSize = edit.size * viewport.width / edit.pageWidth;
                    ctx.save();
                    ctx.fillStyle = edit.color;
                    ctx.font = `${fontSize}px Arial, sans-serif`;
                    ctx.textBaseline = 'top';
                    ctx.fillText(edit.text, edit.x * viewport.width / edit.pageWidth,
                        edit.y * viewport.height / edit.pageHeight);
                    ctx.restore();
                }

                if (edit.type === 'highlight' || edit.type === 'whiteout') {
                    ctx.save();
                    ctx.globalAlpha = edit.type === 'highlight' ? .28 : 1;
                    ctx.fillStyle = edit.type === 'whiteout' ? '#ffffff' : edit.color;
                    ctx.fillRect(
                        edit.x * viewport.width / edit.pageWidth,
                        edit.y * viewport.height / edit.pageHeight,
                        edit.w * viewport.width / edit.pageWidth,
                        edit.h * viewport.height / edit.pageHeight
                    );
                    ctx.restore();
                }

                if (edit.type === 'draw') {
                    if (!edit.points.length) continue;
                    ctx.save();
                    ctx.strokeStyle = edit.color;
                    ctx.lineWidth = edit.size * viewport.width / edit.pageWidth;
                    ctx.lineCap = 'round';
                    ctx.lineJoin = 'round';
                    ctx.beginPath();
                    edit.points.forEach((pt, i) => {
                        const x = pt.x * viewport.width / edit.pageWidth;
                        const y = pt.y * viewport.height / edit.pageHeight;
                        if (i === 0) ctx.moveTo(x, y);
                        else ctx.lineTo(x, y);
                    });
                    ctx.stroke();
                    ctx.restore();
                }
            }
        }

        function editorPoint(e) {
            const canvas = document.getElementById('pdfEditorCanvas');
            const rect = canvas.getBoundingClientRect();
            const sx = canvas.width / rect.width;
            const sy = canvas.height / rect.height;
            return {
                x: (e.clientX - rect.left) * sx,
                y: (e.clientY - rect.top) * sy
            };
        }

        function editorPageCoordinates(point) {
            const canvas = document.getElementById('pdfEditorCanvas');
            return {
                x: point.x / canvas.width,
                y: point.y / canvas.height
            };
        }

        const editorCanvas = document.getElementById('pdfEditorCanvas');

        editorCanvas.addEventListener('mousedown', (e) => {
            if (editorTool === 'select') return;

            const point = editorPoint(e);
            const pageObj = state.pageOrder[editorPageIndex];
            const canvas = editorCanvas;
            const edits = getEditorEdits(pageObj);
            const pageWidth = canvas.width;
            const pageHeight = canvas.height;

            if (editorTool === 'text') {
                const text = prompt("Enter text:");
                if (!text) return;

                const size = Number(document.getElementById('pdfEditorSize').value) * 3 + 10;
                const color = document.getElementById('pdfEditorColor').value;

                edits.push({
                    type: 'text',
                    text,
                    x: point.x,
                    y: point.y,
                    size,
                    color,
                    pageWidth,
                    pageHeight
                });
                renderEditorPage();
                return;
            }

            editorDrawing = true;
            editorStart = point;

            if (editorTool === 'draw') {
                editorCurrentPath = [point];
            }
        });

        editorCanvas.addEventListener('mousemove', (e) => {
            if (!editorDrawing) return;

            const point = editorPoint(e);
            const canvas = editorCanvas;
            const ctx = canvas.getContext('2d');

            if (editorTool === 'draw') {
                const last = editorCurrentPath[editorCurrentPath.length - 1];
                editorCurrentPath.push(point);

                ctx.save();
                ctx.strokeStyle = document.getElementById('pdfEditorColor').value;
                ctx.lineWidth = Number(document.getElementById('pdfEditorSize').value);
                ctx.lineCap = 'round';
                ctx.lineJoin = 'round';
                ctx.beginPath();
                ctx.moveTo(last.x, last.y);
                ctx.lineTo(point.x, point.y);
                ctx.stroke();
                ctx.restore();
                return;
            }

            // Live rectangle preview for highlight/whiteout.
            if (editorTool === 'highlight' || editorTool === 'whiteout') {
                if (editorBaseImage) ctx.putImageData(editorBaseImage, 0, 0);
                const pageObj = state.pageOrder[editorPageIndex];
                drawEditorOverlays(ctx, pageObj, {
                    width: canvas.width,
                    height: canvas.height
                });

                const x = Math.min(editorStart.x, point.x);
                const y = Math.min(editorStart.y, point.y);
                const w = Math.abs(point.x - editorStart.x);
                const h = Math.abs(point.y - editorStart.y);

                ctx.save();
                ctx.globalAlpha = editorTool === 'highlight' ? .28 : 1;
                ctx.fillStyle = editorTool === 'whiteout' ? '#ffffff' :
                    document.getElementById('pdfEditorColor').value;
                ctx.fillRect(x, y, w, h);
                ctx.restore();
            }
        });

        function finishEditorPointer(e) {
            if (!editorDrawing) return;
            editorDrawing = false;

            const pageObj = state.pageOrder[editorPageIndex];
            const edits = getEditorEdits(pageObj);
            const canvas = editorCanvas;
            const end = editorPoint(e);
            const color = document.getElementById('pdfEditorColor').value;
            const size = Number(document.getElementById('pdfEditorSize').value);

            if (editorTool === 'draw' && editorCurrentPath.length > 1) {
                edits.push({
                    type: 'draw',
                    points: editorCurrentPath,
                    color,
                    size,
                    pageWidth: canvas.width,
                    pageHeight: canvas.height
                });
            }

            if (editorTool === 'highlight' || editorTool === 'whiteout') {
                const x = Math.min(editorStart.x, end.x);
                const y = Math.min(editorStart.y, end.y);
                const w = Math.abs(end.x - editorStart.x);
                const h = Math.abs(end.y - editorStart.y);

                if (w > 2 && h > 2) {
                    edits.push({
                        type: editorTool,
                        x, y, w, h,
                        color,
                        pageWidth: canvas.width,
                        pageHeight: canvas.height
                    });
                }
            }

            editorStart = null;
            editorCurrentPath = [];
            renderEditorPage();
        }

        editorCanvas.addEventListener('mouseup', finishEditorPointer);
        editorCanvas.addEventListener('mouseleave', (e) => {
            if (editorDrawing) finishEditorPointer(e);
        });

        function undoEditorEdit() {
            const pageObj = state.pageOrder[editorPageIndex];
            const edits = getEditorEdits(pageObj);
            if (edits.length) {
                edits.pop();
                renderEditorPage();
            }
        }

        function clearEditorPage() {
            const pageObj = state.pageOrder[editorPageIndex];
            const edits = getEditorEdits(pageObj);
            if (!edits.length) return;
            if (confirm("Remove all edits from this page?")) {
                edits.length = 0;
                renderEditorPage();
            }
        }

        function hexToRgb(hex) {
            const value = hex.replace('#', '');
            const n = parseInt(value, 16);
            return {
                r: (n >> 16) & 255,
                g: (n >> 8) & 255,
                b: n & 255
            };
        }

        async function applyPdfEdits(pdfPage, pageObj) {
            const edits = pdfEdits.get(getEditorKey(pageObj));
            if (!edits || !edits.length) return;

            const pageWidth = pdfPage.getWidth();
            const pageHeight = pdfPage.getHeight();

            for (const edit of edits) {
                const sx = pageWidth / edit.pageWidth;
                const sy = pageHeight / edit.pageHeight;

                if (edit.type === 'text') {
                    const rgb = hexToRgb(edit.color);
                    const fontSize = edit.size * sx;
                    pdfPage.drawText(edit.text, {
                        x: edit.x * sx,
                        y: pageHeight - (edit.y * sy) - fontSize,
                        size: fontSize,
                        color: PDFLib.rgb(rgb.r / 255, rgb.g / 255, rgb.b / 255)
                    });
                }

                if (edit.type === 'highlight' || edit.type === 'whiteout') {
                    const rgb = edit.type === 'whiteout'
                        ? { r: 255, g: 255, b: 255 }
                        : hexToRgb(edit.color);

                    pdfPage.drawRectangle({
                        x: edit.x * sx,
                        y: pageHeight - (edit.y + edit.h) * sy,
                        width: edit.w * sx,
                        height: edit.h * sy,
                        color: PDFLib.rgb(rgb.r / 255, rgb.g / 255, rgb.b / 255),
                        opacity: edit.type === 'highlight' ? .28 : 1
                    });
                }

                if (edit.type === 'draw' && edit.points.length > 1) {
                    const rgb = hexToRgb(edit.color);
                    const lineWidth = Math.max(1, edit.size * sx);

                    for (let i = 1; i < edit.points.length; i++) {
                        const a = edit.points[i - 1];
                        const b = edit.points[i];

                        pdfPage.drawLine({
                            start: {
                                x: a.x * sx,
                                y: pageHeight - a.y * sy
                            },
                            end: {
                                x: b.x * sx,
                                y: pageHeight - b.y * sy
                            },
                            thickness: lineWidth,
                            color: PDFLib.rgb(rgb.r / 255, rgb.g / 255, rgb.b / 255),
                            lineCap: PDFLib.LineCapStyle?.Round
                        });
                    }
                }
            }
        }

        window.addEventListener('mouseup', () => isDragging = false);
    </script>
</body>

</html>