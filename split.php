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

        #loadingOverlay {
            display: none;
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(255, 255, 255, 0.85);
            backdrop-filter: blur(6px);
            z-index: 9999;
        }

        .loading-box {
            position: absolute;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            text-align: center;
            background: white;
            padding: 45px 50px;
            border-radius: 20px;
            box-shadow: 0 15px 40px rgba(0, 0, 0, 0.12);
        }

        @keyframes thumbFadeIn {
            from {
                opacity: 0;
            }

            to {
                opacity: 1;
            }
        }

        @keyframes cardPopIn {
            from {
                opacity: 0;
                transform: translateY(10px) scale(0.95);
            }

            to {
                opacity: 1;
                transform: translateY(0) scale(1);
            }
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
            border-radius: 20px;
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

        .btn-prev {
            background: #fff;
            color: var(--primary, #1e293b);
            border: 1.5px dashed var(--primary, #cbd5e1);
            border-radius: 20px;
        }

        .btn-prev:hover {
            background: var(--bg);
            border-style: solid;
        }

        #previewAllBtn,
        #downloadAllBtn,
        #clearAllBtn {
            animation:
                btnAppear 0.35s ease,
                btnFloat 3s ease-in-out infinite 0.35s;
        }

        @keyframes btnAppear {
            from {
                opacity: 0;
                transform: translateY(10px) scale(0.9);
            }

            to {
                opacity: 1;
                transform: translateY(0) scale(1);
            }
        }

        /* Preview All modal */
        .preview-all-grid {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 32px;
            max-height: 78vh;
            overflow-y: auto;
            padding: 4px;
        }

        .preview-all-card {
            display: flex;
            flex-direction: column;
            align-items: center;
            text-align: center;
        }

        .preview-all-thumb-wrap {
            width: 100%;
            aspect-ratio: 3 / 4;
            background: #f1f5f9;
            border: 1px solid #e2e8f0;
            border-radius: 12px;
            overflow: hidden;
            display: flex;
            align-items: center;
            justify-content: center;
        }

        .preview-all-thumb-wrap img {
            max-width: 90%;
            max-height: 90%;
            object-fit: contain;
        }

        .preview-all-name {
            margin-top: 12px;
            font-size: 15px;
            font-weight: 600;
            color: #1e293b;
            width: 100%;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }

        .preview-all-count {
            font-size: 12.5px;
            color: #94a3b8;
            margin-top: 3px;
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
        }

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

        /* Page number badge (text is now set directly by JS, per split group) */

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
            <button class="btn btn-prev" id="previewAllBtn" onclick="showPreviewAll()" style="display:none;">Preview
                All</button>
            <button class="btn btn-main" id="downloadAllBtn" onclick="exportPDF()" style="display:none;">
                Download All
            </button>

            <button class="btn btn-clear" id="clearAllBtn" onclick="confirmClearAll()" style="display:none;">
                Clear All
            </button>
        </div>
    </div>

    <div id="workspace" class="workspace-grid">
        <div class="drop-hint" id="drop-hint"><i class="fa-solid fa-file-arrow-up"></i>Drag and Drop files here</div>
    </div>

    <div id="previewAllModal" class="modal-overlay">
        <div
            style="background: white; padding: 28px; border-radius: 25px; max-width: 1300px; width: 95%; max-height: 90vh;">
            <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:18px;">
                <h3 style="margin:0;">Preview All Files</h3>
                <i class="fa fa-times" style="cursor:pointer; font-size:18px; color:#94a3b8;"
                    onclick="closePreviewAll()"></i>
            </div>
            <div id="previewAllBody" class="preview-all-grid"></div>
        </div>
    </div>

    <div id="customAlert" class="modal-overlay">
        <div
            style="background: white; padding: 32px; border-radius: 25px; text-align: center; max-width: 360px; width: 90%;">
            <h3 id="alertTitle">Status</h3>
            <p id="alertMessage"></p>
            <button class="btn btn-main" id="alertBtn" onclick="closeAlert()">OK</button>
        </div>
    </div>

    <div id="loadingOverlay">
        <div class="loading-box">
            <i class="fa fa-spinner fa-spin"
                style="font-size:28px; color:var(--primary, #1e293b); margin-bottom:14px; display:inline-block;"></i>
            <p id="loadingTitle" style="margin:0; font-weight:bold; color:#333;">Processing...</p>
            <p style="margin:10px 0 0; font-size:13px; color:#999;">Please wait...</p>
        </div>
    </div>

    <div id="confirmOverlay"
        style="display:none; position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(255, 255, 255, 0.85); backdrop-filter:blur(6px); z-index:9999; align-items:center; justify-content:center;">
        <div class="loading-box" style="position:static; transform:none; max-width:320px; width:90%;">
            <i class="fa-solid fa-triangle-exclamation"
                style="font-size:50px; color:#ef4444; margin-bottom:20px; display:inline-block;"></i>
            <p style="margin:0; font-weight:bold; color:#ef4444; margin-bottom:20px;">Clear All Files?</p>
            <div style="display:flex; gap:20px; justify-content:center;">
                <button class="btn btn-clear" style="flex:1; justify-content:center;"
                    onclick="closeConfirmClearAll()">Cancel</button>
                <button class="btn" style="flex:1; justify-content:center; background:#ef4444; color:#fff;"
                    onclick="location.reload()">Yes</button>
            </div>
        </div>
    </div>

    <div id="homeConfirmOverlay"
        style="display:none; position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(255, 255, 255, 0.85); backdrop-filter:blur(6px); z-index:9999; align-items:center; justify-content:center;">
        <div class="loading-box" style="position:static; transform:none; max-width:320px; width:90%;">
            <i class="fa-solid fa-triangle-exclamation"
                style="font-size:50px; color:#ef4444; margin-bottom:20px; display:inline-block;"></i>
            <p style="margin:0; font-weight:bold; color:#ef4444; margin-bottom:20px;">Leave this page?</p>
            <p style="margin:0 0 20px; font-size:13px; color:#666;">Your uploaded files will be lost.</p>
            <div style="display:flex; gap:20px; justify-content:center;">
                <button class="btn btn-clear" style="flex:1; justify-content:center;"
                    onclick="closeHomeConfirm()">Cancel</button>
                <button class="btn" style="flex:1; justify-content:center; background:#ef4444; color:#fff;"
                    onclick="window.location.href='index.html'">Yes</button>
            </div>
        </div>
    </div>

    <div id="previewModal" class="modal-overlay" onclick="closePreview()">
        <button class="nav-arrow" id="prevArrow" onclick="navigatePreview(-1, event)"><i
                class="fa fa-chevron-left"></i></button>
        <img id="previewImage" src="" onclick="event.stopPropagation()">
        <button class="nav-arrow" id="nextArrow" onclick="navigatePreview(1, event)"><i
                class="fa fa-chevron-right"></i></button>
    </div>

    <a href="index.html" class="home-btn" title="Back to Home" onclick="return confirmGoHome()"><i
            class="fa fa-home"></i></a>

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
            return 'page_' + Date.now() + '_' + Math.random().toString(36).substring(2, 9);
        }

        function cloneBuffer(buffer) {
            const dst = new ArrayBuffer(buffer.byteLength);
            new Uint8Array(dst).set(new Uint8Array(buffer));
            return dst;
        }

        document.getElementById('file-selector').addEventListener('change', (e) => {
            handleFiles(Array.from(e.target.files));
        });

        let externalDragOverCard = null;
        let externalDropSide = null;

        workspace.addEventListener('dragover', (e) => {
            e.preventDefault();

            // Only show insertion indicator for files dragged from
            // Windows File Explorer.
            const hasFiles = e.dataTransfer &&
                Array.from(e.dataTransfer.types || []).includes('Files');

            if (!hasFiles) {
                workspace.classList.add('drag-active');
                return;
            }

            workspace.classList.add('drag-active');

            const target = e.target instanceof Element
                ? e.target
                : null;

            const card = target
                ? target.closest('.page-card')
                : null;

            // Clear previous indicator
            if (externalDragOverCard && externalDragOverCard !== card) {
                externalDragOverCard.style.borderLeft = '';
                externalDragOverCard.style.borderRight = '';
            }

            externalDragOverCard = card;
            externalDropSide = null;

            if (!card) {
                return;
            }

            const rect = card.getBoundingClientRect();
            const middle = rect.left + (rect.width / 2);

            if (e.clientX < middle) {
                card.style.borderLeft = '5px solid var(--primary)';
                card.style.borderRight = '';
                externalDropSide = 'before';
            } else {
                card.style.borderLeft = '';
                card.style.borderRight = '5px solid var(--primary)';
                externalDropSide = 'after';
            }

            // Important for Firefox/Chrome external file dragging
            e.dataTransfer.dropEffect = 'copy';
        });

        workspace.addEventListener('dragleave', (e) => {
            // Ignore dragleave events when moving between
            // children inside the workspace.
            if (e.relatedTarget && workspace.contains(e.relatedTarget)) {
                return;
            }

            workspace.classList.remove('drag-active');

            if (externalDragOverCard) {
                externalDragOverCard.style.borderLeft = '';
                externalDragOverCard.style.borderRight = '';
            }

            externalDragOverCard = null;
            externalDropSide = null;
        });

        workspace.addEventListener('drop', (e) => {
            e.preventDefault();
            e.stopPropagation();

            workspace.classList.remove('drag-active');

            // Clear visual insertion indicator
            workspace.querySelectorAll('.page-card').forEach(card => {
                card.style.borderLeft = '';
                card.style.borderRight = '';
            });

            externalDragOverCard = null;

            const files = Array.from(e.dataTransfer.files || []);

            if (files.length === 0) {
                externalDropSide = null;
                return;
            }

            // Find the page card underneath the mouse.
            const target = e.target instanceof Element
                ? e.target
                : null;

            const targetCard = target
                ? target.closest('.page-card')
                : null;

            // No page card = append to the end.
            if (!targetCard) {
                externalDropSide = null;
                handleFiles(files, state.pageOrder.length);
                return;
            }

            const cards = Array.from(
                workspace.querySelectorAll('.page-card')
            );

            const targetIndex = cards.indexOf(targetCard);

            if (targetIndex === -1) {
                externalDropSide = null;
                handleFiles(files, state.pageOrder.length);
                return;
            }

            const rect = targetCard.getBoundingClientRect();

            // Use the actual mouse position at drop time.
            const middle = rect.left + (rect.width / 2);

            const insertIndex =
                e.clientX < middle
                    ? targetIndex
                    : targetIndex + 1;

            externalDropSide = null;

            handleFiles(files, insertIndex);
        });

        document.getElementById('file-selector')
            .addEventListener('change', (e) => {

                handleFiles(
                    Array.from(e.target.files)
                );

            });

        async function handleFiles(files, insertIndex = -1) {
            if (files.length === 0) return;
            showAlert("Processing...", false);

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
                            id: generateFileId(),
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
            const previewBtn = document.getElementById('previewAllBtn');
            const downloadBtn = document.getElementById('downloadAllBtn');
            const clearBtn = document.getElementById('clearAllBtn');

            if (state.pageOrder.length === 0) {
                workspace.innerHTML = '<div class="drop-hint" id="drop-hint"><i class="fa-solid fa-file-arrow-up"></i>Drag and Drop files here</div>';

                if (previewBtn) previewBtn.style.display = 'none';
                if (downloadBtn) downloadBtn.style.display = 'none';
                if (clearBtn) clearBtn.style.display = 'none';

                return;
            }

            if (previewBtn) previewBtn.style.display = 'inline-flex';
            if (downloadBtn) downloadBtn.style.display = 'inline-flex';
            if (clearBtn) clearBtn.style.display = 'inline-flex';

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
            renumberSegments();
        }

        function renumberSegments() {
            const visibleHeaders = Array.from(workspace.querySelectorAll('.segment-header')).filter(h =>
                h.classList.contains('force-show') || h.previousElementSibling.classList.contains('split-active')
            );
            visibleHeaders.forEach((h, i) => {
                const input = h.querySelector('.rename-input');
                if (input) input.placeholder = `File ${i + 1}`;
            });

            // Badge numbers restart at #1 for every split group, instead of
            // counting straight through the whole document.
            let posInGroup = 0;
            Array.from(workspace.querySelectorAll('.page-card')).forEach((card) => {
                posInGroup++;
                const badge = card.querySelector('.badge');
                if (badge) badge.textContent = `#${posInGroup}`;
                if (card.classList.contains('split-active')) posInGroup = 0;
            });
        }

        function createRenameBar(idx) {
            const div = document.createElement('div');
            div.className = 'segment-header';
            div.dataset.forIdx = idx;

            div.innerHTML = `
        <span class="segment-label">File Name:</span>
        <input type="text" class="rename-input" value="${state.segmentNames[idx + 1] || ''}" placeholder="File" oninput="state.segmentNames[${idx + 1}] = this.value">
        <button class="btn btn-main" style="height:34px; font-size:12px;" onclick="downloadSingleGroupFromDOM(${idx})">Download</button>
    `;
            return div;
        }

        function createPageCard(pageObj, index) {
            const card = document.createElement('div');
            card.className = 'page-card';
            card.draggable = true;
            if (state.selectedIndices.has(index)) card.classList.add('selected');
            if (state.splits.has(pageObj.id)) {
                card.classList.add('split-active');
            }

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

                if (state.splits.has(pageObj.id)) {
                    state.splits.delete(pageObj.id);
                    card.classList.remove('split-active');
                } else {
                    state.splits.add(pageObj.id);
                    card.classList.add('split-active');
                }

                renumberSegments();
            };

            card.querySelector('.rotate-btn').onclick = (e) => {
                e.stopPropagation();
                pageObj.rotation = (pageObj.rotation + 90) % 360;
                card.querySelector('canvas').style.transform = `rotate(${pageObj.rotation}deg)`;
            };

            card.querySelector('.delete-btn').onclick = (e) => {
                e.stopPropagation();
                const removedPage = state.pageOrder[index];

                if (removedPage) {
                    state.splits.delete(removedPage.id);
                }

                state.pageOrder.splice(index, 1);
                state.selectedIndices.clear();
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
                e.stopPropagation();

                card.style.borderLeft = "";
                card.style.borderRight = "";

                const rect = card.getBoundingClientRect();
                const relX = e.clientX - rect.left;

                const dropInsertIndex =
                    relX < rect.width / 2
                        ? index
                        : index + 1;

                // File dragged from Windows/File Explorer
                if (e.dataTransfer.files.length > 0) {
                    handleFiles(
                        Array.from(e.dataTransfer.files),
                        dropInsertIndex
                    );
                    return;
                }

                // Existing page reordering
                const sortedSelected = Array.from(state.selectedIndices)
                    .sort((a, b) => a - b);

                if (sortedSelected.length === 0) return;

                const itemsToMove = sortedSelected.map(
                    i => state.pageOrder[i]
                );

                for (let i = sortedSelected.length - 1; i >= 0; i--) {
                    state.pageOrder.splice(sortedSelected[i], 1);
                }

                let finalInsert = dropInsertIndex;

                const shift = sortedSelected.filter(
                    i => i < dropInsertIndex
                ).length;

                finalInsert -= shift;

                if (finalInsert < 0) {
                    finalInsert = 0;
                }

                if (finalInsert > state.pageOrder.length) {
                    finalInsert = state.pageOrder.length;
                }

                state.pageOrder.splice(
                    finalInsert,
                    0,
                    ...itemsToMove
                );

                state.selectedIndices.clear();

                renderWorkspace();
            };

            return card;
        }

        async function drawThumb(pageObj, canvasId) {
            try {
                const source = sourcePdfs.get(pageObj.fileId);
                const page = await source.pdfjsDoc.getPage(pageObj.originalIdx + 1);
                const viewport = page.getViewport({
                    scale: 3
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

        function splitIntoGroups() {
            const groups = [];
            let currentGroup = [];

            state.pageOrder.forEach((pageObj, index) => {
                currentGroup.push(pageObj);

                if (
                    state.splits.has(pageObj.id) &&
                    index < state.pageOrder.length - 1
                ) {
                    groups.push(currentGroup);
                    currentGroup = [];
                }
            });

            if (currentGroup.length > 0) {
                groups.push(currentGroup);
            }

            return groups;
        }

        function showPreviewAll() {
            if (state.pageOrder.length === 0) return;

            const cards = Array.from(workspace.querySelectorAll('.page-card'));
            const groups = [];
            let currentGroup = [];
            let groupStart = 0;
            state.pageOrder.forEach((pageObj) => {
                currentGroup.push(pageObj);

                if (state.splits.has(pageObj.id)) {
                    groups.push({
                        pages: currentGroup
                    });
                    currentGroup = [];
                }
            });
            if (currentGroup.length > 0) {
                groups.push({ pages: currentGroup, firstCardIndex: groupStart });
            }

            const headers = Array.from(workspace.querySelectorAll('.segment-header')).filter(h =>
                h.classList.contains('force-show') || h.previousElementSibling.classList.contains('split-active')
            );

            const body = document.getElementById('previewAllBody');
            body.innerHTML = groups.map((group, i) => {
                const nameInput = headers[i] ? headers[i].querySelector('.rename-input') : null;
                const name = sanitizeFileName((nameInput && nameInput.value) || (nameInput && nameInput.placeholder) || `File ${i + 1}`);
                const pageCount = group.pages.length;
                return `
                    <div class="preview-all-card" style="animation: cardPopIn 0.35s ease forwards; animation-delay: ${i * 45}ms; opacity: 0;">
                        <div class="preview-all-thumb-wrap" id="preview-thumb-${i}">
                            <i class="fa fa-spinner fa-spin" style="color:#cbd5e1; font-size:22px;"></i>
                        </div>
                        <div class="preview-all-name">${name}</div>
                        <div class="preview-all-count">${pageCount} page${pageCount > 1 ? 's' : ''}</div>
                    </div>
                `;
            }).join('');

            document.getElementById('previewAllModal').style.display = 'flex';
            const previewModalContent = document.getElementById('previewAllModal').querySelector('div');
            previewModalContent.classList.remove('modal-closing');
            previewModalContent.classList.add('modal-animating');

            groups.forEach(async (group, i) => {
                const firstPage = group.pages[0];
                const wrap = document.getElementById(`preview-thumb-${i}`);
                if (!firstPage || !wrap) return;
                try {
                    const source = sourcePdfs.get(firstPage.fileId);
                    const page = await source.pdfjsDoc.getPage(firstPage.originalIdx + 1);
                    const viewport = page.getViewport({ scale: 10 });
                    const tempCanvas = document.createElement('canvas');
                    tempCanvas.width = viewport.width;
                    tempCanvas.height = viewport.height;
                    await page.render({ canvasContext: tempCanvas.getContext('2d'), viewport }).promise;

                    const img = document.createElement('img');
                    img.src = tempCanvas.toDataURL('image/png');
                    img.style.transform = `rotate(${firstPage.rotation}deg)`;
                    img.style.animation = 'thumbFadeIn 0.35s ease forwards';
                    wrap.innerHTML = '';
                    wrap.appendChild(img);
                } catch (e) {
                    console.error(e);
                    wrap.innerHTML = '<i class="fa fa-file" style="color:#cbd5e1; font-size:22px;"></i>';
                }
            });
        }

        function closePreviewAll() {
            const previewModal = document.getElementById('previewAllModal');
            const previewModalContent = previewModal.querySelector('div');

            previewModalContent.classList.remove('modal-animating');
            previewModalContent.classList.add('modal-closing');

            setTimeout(() => {
                previewModal.style.display = 'none';
                previewModalContent.classList.remove('modal-closing');
            }, 200);
        }

        function sanitizeFileName(name) {
            return name.replace(/[\/\\:*?"<>|]/g, '');
        }

        async function downloadSingleGroupFromDOM(headerIdx) {
            showAlert("Preparing download...", false);
            const groups = splitIntoGroups();
            const headers = Array.from(workspace.querySelectorAll('.segment-header')).filter(h =>
                h.classList.contains('force-show') || h.previousElementSibling.classList.contains('split-active')
            );
            const groupIdx = headers.findIndex(h => h.dataset.forIdx == headerIdx || (headerIdx === -1 && h.classList.contains('force-show')));

            const bytes = await generatePdfBlob(groups[groupIdx]);
            const nameInput = headers[groupIdx].querySelector('.rename-input');
            downloadBlob(bytes, sanitizeFileName(nameInput.value || `File ${groupIdx + 1}`));
            closeAlert();
        }

        async function exportPDF() {
            if (state.pageOrder.length === 0) return showAlert("No pages to export.");
            showAlert("Generating ZIP file...", false);
            const zip = new JSZip();
            const groups = splitIntoGroups();
            const headers = Array.from(workspace.querySelectorAll('.segment-header')).filter(h =>
                h.classList.contains('force-show') || h.previousElementSibling.classList.contains('split-active')
            );

            for (let i = 0; i < groups.length; i++) {
                const bytes = await generatePdfBlob(groups[i]);
                const name = sanitizeFileName(headers[i].querySelector('.rename-input').value || `File ${i + 1}`);
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

        function confirmClearAll() {
            if (state.pageOrder.length === 0) {
                location.reload();
                return;
            }
            const overlay = document.getElementById('confirmOverlay');
            const box = overlay.querySelector('.loading-box');
            overlay.style.display = 'flex';
            box.classList.remove('modal-closing');
            box.classList.add('modal-animating');
        }

        function closeConfirmClearAll() {
            const overlay = document.getElementById('confirmOverlay');
            const box = overlay.querySelector('.loading-box');
            box.classList.remove('modal-animating');
            box.classList.add('modal-closing');
            setTimeout(() => {
                overlay.style.display = 'none';
                box.classList.remove('modal-closing');
            }, 200);
        }

        function confirmGoHome() {
            if (state.pageOrder.length === 0) return true;
            const overlay = document.getElementById('homeConfirmOverlay');
            const box = overlay.querySelector('.loading-box');
            overlay.style.display = 'flex';
            box.classList.remove('modal-closing');
            box.classList.add('modal-animating');
            return false;
        }

        function closeHomeConfirm() {
            const overlay = document.getElementById('homeConfirmOverlay');
            const box = overlay.querySelector('.loading-box');
            box.classList.remove('modal-animating');
            box.classList.add('modal-closing');
            setTimeout(() => {
                overlay.style.display = 'none';
                box.classList.remove('modal-closing');
            }, 200);
        }

        function showAlert(msg, showButton = true) {
            if (!showButton) {
                document.getElementById('loadingTitle').innerText = msg;
                document.getElementById('loadingOverlay').style.display = 'block';
                return;
            }
            document.getElementById('alertTitle').innerText = "Status";
            document.getElementById('alertMessage').innerText = msg;
            document.getElementById('alertBtn').style.display = 'inline-flex';
            document.getElementById('customAlert').style.display = 'flex';
        }

        function closeAlert() {
            document.getElementById('customAlert').style.display = 'none';
            document.getElementById('loadingOverlay').style.display = 'none';
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
        window.addEventListener('mouseup', () => isDragging = false);
    </script>
</body>

</html>
