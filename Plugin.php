<?php

namespace App\Plugins\ExmentTranslator;

use Encore\Admin\Widgets\Box;
use Exceedone\Exment\Services\Plugin\PluginPageBase;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use ZipArchive;
use GuzzleHttp\Client;

class Plugin extends PluginPageBase
{
    protected $useCustomOption = true;
    
    /**
     * プラグインページではCSRF検証を無効化
     */
    public $disableSession = false;

    /**
     * メインページの表示
     */
    public function index()
    {
        // API キーが設定されているか確認
        $hasApiKey = !is_null($this->plugin->getCustomOption('gemini_api_key'));
        $uploadUrl = $this->plugin->getFullUrl('upload');

        $html = $this->generateIndexHtml($hasApiKey, $uploadUrl);

        $box = new Box('AI翻訳ツール', $html);
        return $box;
    }

    /**
     * ファイルアップロード
     */
    public function upload()
    {
        // エラーをJSONで返すようにエラーハンドラを設定
        set_error_handler(function($errno, $errstr, $errfile, $errline) {
            throw new \ErrorException($errstr, 0, $errno, $errfile, $errline);
        });
        
        try {
            $request = request();
            \Log::info('Upload request received');
            
            if (!$request->hasFile('file')) {
                \Log::error('No file in request');
                restore_error_handler();
                return response()->json(['error' => 'ファイルが選択されていません'], 400);
            }

            $file = $request->file('file');
            $originalName = $file->getClientOriginalName();
            \Log::info('File received: ' . $originalName);
            
            $extension = strtolower($file->getClientOriginalExtension());

            if (!in_array($extension, ['xlsx', 'docx'])) {
                restore_error_handler();
                return response()->json(['error' => 'サポートされていないファイル形式です'], 400);
            }

            // 一時ディレクトリ（より確実な場所を使用）
            $tempDir = sys_get_temp_dir() . '/exment_translator';
            if (!is_dir($tempDir)) {
                if (!mkdir($tempDir, 0755, true)) {
                    throw new \Exception('一時ディレクトリの作成に失敗しました');
                }
            }

            // ファイルを一時保存
            $tempFileName = uniqid('trans_') . '.' . $extension;
            $fullPath = $tempDir . '/' . $tempFileName;
            
            if (!$file->move($tempDir, $tempFileName)) {
                throw new \Exception('ファイルの保存に失敗しました');
            }
            
            \Log::info('File saved to: ' . $fullPath);

            // ファイルが正しく保存されたか確認
            if (!file_exists($fullPath) || !is_readable($fullPath)) {
                throw new \Exception('保存されたファイルにアクセスできません');
            }

            // ユニークな文字列を抽出
            $uniqueStrings = [];
            
            if ($extension === 'xlsx') {
                if (!class_exists('PhpOffice\PhpSpreadsheet\IOFactory')) {
                    throw new \Exception('PhpSpreadsheetライブラリが利用できません');
                }
                $uniqueStrings = $this->extractExcelStrings($fullPath);
            } else {
                if (!class_exists('ZipArchive')) {
                    throw new \Exception('ZipArchiveクラスが利用できません');
                }
                $uniqueStrings = $this->extractWordStrings($fullPath);
            }

            \Log::info('Unique strings found: ' . count($uniqueStrings));

            if (empty($uniqueStrings)) {
                throw new \Exception('翻訳可能なテキストが見つかりませんでした');
            }

            // セッションに保存
            $request->session()->put('translation_file_path', $fullPath);
            $request->session()->put('translation_file_type', $extension);
            $request->session()->put('translation_unique_strings', $uniqueStrings);
            $request->session()->put('translation_original_name', $originalName);
            $request->session()->save();

            restore_error_handler();
            
            return response()->json([
                'success' => true,
                'uniqueCount' => count($uniqueStrings),
                'translateUrl' => $this->plugin->getFullUrl('translate'),
            ]);

        } catch (\Throwable $e) {
            restore_error_handler();
            \Log::error('Upload error: ' . $e->getMessage());
            \Log::error('File: ' . $e->getFile() . ' Line: ' . $e->getLine());
            \Log::error($e->getTraceAsString());
            
            return response()->json([
                'error' => $e->getMessage(),
                'file' => basename($e->getFile()),
                'line' => $e->getLine()
            ], 500);
        }
    }

    /**
     * 翻訳実行
     */
    public function translate()
    {
        set_error_handler(function($errno, $errstr, $errfile, $errline) {
            throw new \ErrorException($errstr, 0, $errno, $errfile, $errline);
        });
        
        try {
            $request = request();
            $targetLanguage = $request->input('target_language', '日本語');
            $customInstruction = $request->input('custom_instruction', '');

            // セッションから翻訳対象の文字列を取得
            $uniqueStrings = $request->session()->get('translation_unique_strings');
            $filePath = $request->session()->get('translation_file_path');
            $fileType = $request->session()->get('translation_file_type');
            $originalName = $request->session()->get('translation_original_name');

            if (!$uniqueStrings || !$filePath) {
                restore_error_handler();
                return response()->json(['error' => 'セッションの有効期限が切れました。ファイルを再アップロードしてください。'], 400);
            }
            
            if (!file_exists($filePath)) {
                restore_error_handler();
                return response()->json(['error' => 'アップロードされたファイルが見つかりません。再アップロードしてください。'], 400);
            }

            // カスタム設定からAPIキーを取得
            $apiKey = $this->plugin->getCustomOption('gemini_api_key');
            if (empty($apiKey)) {
                restore_error_handler();
                return response()->json(['error' => 'Gemini APIキーが設定されていません。プラグイン設定画面で設定してください。'], 400);
            }

            // 翻訳を実行
            $translations = $this->translateStrings($uniqueStrings, $targetLanguage, $customInstruction, $apiKey);

            // 翻訳結果をファイルに適用
            $tempDir = sys_get_temp_dir() . '/exment_translator';
            if (!is_dir($tempDir)) {
                mkdir($tempDir, 0755, true);
            }
            
            if ($fileType === 'xlsx') {
                $outputPath = $this->applyExcelTranslations($filePath, $translations, $tempDir);
            } else {
                $outputPath = $this->applyWordTranslations($filePath, $translations, $tempDir);
            }

            // 安全なファイル名を生成（ASCII文字のみ）
            $safeFileName = uniqid('translated_') . '.' . $fileType;
            $finalPath = $tempDir . '/' . $safeFileName;
            
            if (file_exists($finalPath)) {
                @unlink($finalPath);
            }
            rename($outputPath, $finalPath);
            
            \Log::info('File saved to: ' . $finalPath);
            \Log::info('File exists check: ' . (file_exists($finalPath) ? 'YES' : 'NO'));

            // 元のファイルを削除
            if (file_exists($filePath)) {
                @unlink($filePath);
            }
            
            // オリジナルのファイル名をセッションに保存
            $outputFileName = pathinfo($originalName, PATHINFO_FILENAME) . '_translated.' . $fileType;
            $request->session()->put('download_filename', $safeFileName);
            $request->session()->put('download_original_name', $outputFileName);
            $request->session()->save();

            restore_error_handler();
            
            return response()->json([
                'success' => true,
                'downloadUrl' => $this->plugin->getFullUrl('download/' . $safeFileName),
            ]);

        } catch (\Throwable $e) {
            restore_error_handler();
            \Log::error('Translation error: ' . $e->getMessage());
            \Log::error('File: ' . $e->getFile() . ' Line: ' . $e->getLine());
            \Log::error($e->getTraceAsString());
            
            // GuzzleHTTPの例外からHTTPステータスコードを取得
            $errorMessage = $e->getMessage();
            $statusCode = 500;
            
            if (method_exists($e, 'getResponse') && $e->getResponse()) {
                $statusCode = $e->getResponse()->getStatusCode();
            }
            
            // 429エラー（レート制限）の場合は分かりやすいメッセージを返す
            if ($statusCode === 429 || strpos($errorMessage, '429') !== false || strpos($errorMessage, 'quota') !== false) {
                return response()->json([
                    'error' => 'Google Gemini APIの利用レートの制限に達しました。対処方法:1. 数分待ってから再度お試しください. 別のAPIキーを使用してください. Google AI Studioで課金プランをご確認ください https://aistudio.google.com/',
                    'error_type' => 'rate_limit'
                ], 429);
            }
            
            return response()->json([
                'error' => '翻訳処理でエラーが発生しました:' . $errorMessage,
                'file' => basename($e->getFile()),
                'line' => $e->getLine()
            ], 500);
        }
    }

    /**
     * ダウンロード
     */
    public function download($filename)
    {
        try {
            // ファイル名のサニタイズ
            $filename = basename($filename);
            
            $tempDir = sys_get_temp_dir() . '/exment_translator';
            $filePath = $tempDir . '/' . $filename;

            \Log::info('Download request for: ' . $filename);
            \Log::info('Looking for file at: ' . $filePath);
            \Log::info('File exists: ' . (file_exists($filePath) ? 'YES' : 'NO'));

            if (!file_exists($filePath)) {
                \Log::error('Download file not found: ' . $filePath);
                // ディレクトリ内のファイル一覧をログに記録
                if (is_dir($tempDir)) {
                    $files = scandir($tempDir);
                    \Log::info('Files in directory: ' . implode(', ', $files));
                }
                abort(404, 'ファイルが見つかりません');
            }
            
            if (!is_readable($filePath)) {
                \Log::error('Download file not readable: ' . $filePath);
                abort(403, 'ファイルにアクセスできません');
            }

            // セッションからオリジナルのファイル名を取得
            $request = request();
            $originalName = $request->session()->get('download_original_name', $filename);
            \Log::info('Original filename from session: ' . $originalName);
            
            // ファイルの拡張子に応じたMIMEタイプを設定
            $extension = pathinfo($filename, PATHINFO_EXTENSION);
            $mimeType = $extension === 'xlsx' 
                ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

            // ファイルをバイナリモードで読み込み
            $fileContent = file_get_contents($filePath);
            
            if ($fileContent === false) {
                \Log::error('Failed to read file: ' . $filePath);
                abort(500, 'ファイルの読み込みに失敗しました');
            }

            // 強制ダウンロード用のヘッダーを設定（オリジナル名を使用）
            $headers = [
                'Content-Type' => $mimeType,
                'Content-Disposition' => 'attachment; filename="' . $originalName . '"',
                'Content-Length' => strlen($fileContent),
                'Cache-Control' => 'no-cache, no-store, must-revalidate',
                'Pragma' => 'no-cache',
                'Expires' => '0',
            ];

            // ファイルを削除
            @unlink($filePath);

            return response($fileContent, 200, $headers);
            
        } catch (\Throwable $e) {
            \Log::error('Download error: ' . $e->getMessage());
            abort(500, 'ダウンロード処理でエラーが発生しました');
        }
    }

    /**
     * Excelから文字列を抽出
     */
    private function extractExcelStrings($filePath)
    {
        $spreadsheet = IOFactory::load($filePath);
        $uniqueStrings = [];

        foreach ($spreadsheet->getAllSheets() as $sheet) {
            foreach ($sheet->getRowIterator() as $row) {
                foreach ($row->getCellIterator() as $cell) {
                    $value = $cell->getValue();
                    if (is_string($value) && trim($value) !== '' && !is_numeric($value)) {
                        $uniqueStrings[$value] = true;
                    }
                }
            }
        }

        return array_keys($uniqueStrings);
    }

    /**
     * Wordから文字列を抽出
     */
    private function extractWordStrings($filePath)
    {
        $zip = new ZipArchive();
        $zip->open($filePath);
        
        $xmlContent = $zip->getFromName('word/document.xml');
        if ($xmlContent === false) {
            throw new \Exception('Word文書の解析に失敗しました');
        }

        $uniqueStrings = [];
        $xml = simplexml_load_string($xmlContent);
        $xml->registerXPathNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main');
        
        $texts = $xml->xpath('//w:t');
        foreach ($texts as $text) {
            $value = (string)$text;
            if (trim($value) !== '' && !is_numeric($value)) {
                $uniqueStrings[$value] = true;
            }
        }

        $zip->close();
        return array_keys($uniqueStrings);
    }

    /**
     * 文字列を翻訳
     */
    private function translateStrings($strings, $targetLanguage, $customInstruction)
    {
        $apiKey = $this->plugin->getCustomOption('gemini_api_key');
        if (!$apiKey) {
            throw new \Exception('Gemini APIキーが設定されていません');
        }

        $client = new Client();
        $translations = [];
        
        // バッチ処理（10件ずつに削減してレート制限を回避）
        $batches = array_chunk($strings, 10);
        $batchCount = count($batches);
        
        foreach ($batches as $index => $batch) {
            // 2回目以降のリクエストの前に待機（レート制限対策）
            if ($index > 0) {
                \Log::info("Waiting 2 seconds before next batch (batch " . ($index + 1) . "/{$batchCount})...");
                sleep(2); // 2秒待機
            }
            
            $prompt = $this->buildTranslationPrompt($batch, $targetLanguage, $customInstruction);
            
            try {
                $response = $client->post('https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent', [
                    'headers' => [
                        'Content-Type' => 'application/json',
                        'x-goog-api-key' => $apiKey,
                    ],
                    'json' => [
                        'contents' => [
                            [
                                'parts' => [
                                    ['text' => $prompt]
                                ]
                            ]
                        ],
                        'generationConfig' => [
                            'response_mime_type' => 'application/json',
                        ],
                    ],
                    'timeout' => 60, // タイムアウトを60秒に延長
                    'connect_timeout' => 10, // 接続タイムアウトは10秒
                    'http_errors' => true, // HTTPエラーで例外を投げる
                ]);

                $result = json_decode($response->getBody()->getContents(), true);
                $translatedText = $result['candidates'][0]['content']['parts'][0]['text'] ?? '[]';
                $translatedBatch = json_decode($translatedText, true);

                if (is_array($translatedBatch) && count($translatedBatch) === count($batch)) {
                    $translations = array_merge($translations, array_combine($batch, $translatedBatch));
                } else {
                    // フォールバック: 1:1マッピング
                    foreach ($batch as $text) {
                        $translations[$text] = $text;
                    }
                }
                
                \Log::info("Batch " . ($index + 1) . "/{$batchCount} completed successfully");
                
            } catch (\GuzzleHttp\Exception\RequestException $e) {
                \Log::error("Batch " . ($index + 1) . "/{$batchCount} failed: " . $e->getMessage());
                
                // レスポンスボディを取得して詳細なエラー情報をログに記録
                if ($e->hasResponse()) {
                    $statusCode = $e->getResponse()->getStatusCode();
                    $responseBody = $e->getResponse()->getBody()->getContents();
                    \Log::error("HTTP Status: {$statusCode}, Response: {$responseBody}");
                }
                
                // 429エラーの場合は即座に再スロー
                if (strpos($e->getMessage(), '429') !== false || 
                    ($e->hasResponse() && $e->getResponse()->getStatusCode() === 429)) {
                    throw $e;
                }
                
                // その他のエラーの場合は元のテキストをそのまま使用
                foreach ($batch as $text) {
                    $translations[$text] = $text;
                }
            } catch (\Exception $e) {
                \Log::error("Batch " . ($index + 1) . "/{$batchCount} unexpected error: " . $e->getMessage());
                // 元のテキストをそのまま使用
                foreach ($batch as $text) {
                    $translations[$text] = $text;
                }
            }
        }

        return $translations;
    }

    /**
     * 翻訳プロンプトを構築（Node.js版と同じ形式）
     */
    private function buildTranslationPrompt($texts, $targetLanguage, $customInstruction)
    {
        $customPart = $customInstruction ? "4. Custom Instruction from user: {$customInstruction}\n" : '';
        
        return "You are a professional translator.\n"
             . "Translate the following array of text strings into {$targetLanguage}.\n\n"
             . "Rules:\n"
             . "1. Maintain the exact order of the input array.\n"
             . "2. Preserve any special formatting codes, numbers, or symbols.\n"
             . "3. If a string is a proper noun or code that should not be translated, keep it as is.\n"
             . $customPart
             . "5. Return ONLY the JSON array of strings.\n\n"
             . "Input Array:\n" . json_encode($texts, JSON_UNESCAPED_UNICODE);
    }

    /**
     * Excelに翻訳を適用
     */
    private function applyExcelTranslations($filePath, $translations, $tempDir)
    {
        // 元のファイルをコピーして、そのコピーに翻訳を適用
        $outputPath = $tempDir . '/' . uniqid('excel_') . '.xlsx';
        if (!copy($filePath, $outputPath)) {
            throw new \Exception('Excelファイルのコピーに失敗しました');
        }
        
        // コピーしたファイルを読み込んで編集
        $spreadsheet = IOFactory::load($outputPath);

        foreach ($spreadsheet->getAllSheets() as $sheet) {
            $sheet->getCell('A1'); // シートをアクティブ化
            $highestRow = $sheet->getHighestRow();
            $highestColumn = $sheet->getHighestColumn();
            
            // 範囲を指定して処理（メモリ効率化）
            for ($row = 1; $row <= $highestRow; $row++) {
                for ($col = 'A'; $col <= $highestColumn; $col++) {
                    $cell = $sheet->getCell($col . $row);
                    $value = $cell->getValue();
                    
                    // 文字列かつ翻訳が存在する場合のみ適用
                    if (is_string($value) && isset($translations[$value])) {
                        $cell->setValueExplicit(
                            $translations[$value],
                            \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING
                        );
                    }
                }
            }
        }
        
        // 既存のファイルを上書き保存
        $writer = new Xlsx($spreadsheet);
        $writer->setPreCalculateFormulas(false);
        $writer->save($outputPath);
        
        // メモリ解放
        $spreadsheet->disconnectWorksheets();
        unset($spreadsheet);

        return $outputPath;
    }

    /**
     * Wordに翻訳を適用
     */
    private function applyWordTranslations($filePath, $translations, $tempDir)
    {
        // 元のファイルをコピー
        $outputPath = $tempDir . '/' . uniqid('word_') . '.docx';
        if (!copy($filePath, $outputPath)) {
            throw new \Exception('Word文書のコピーに失敗しました');
        }
        
        $zip = new ZipArchive();
        if ($zip->open($outputPath) !== true) {
            throw new \Exception('Word文書を開けませんでした');
        }
        
        $xmlContent = $zip->getFromName('word/document.xml');
        if ($xmlContent === false) {
            $zip->close();
            throw new \Exception('Word文書のXMLが見つかりません');
        }
        
        // 元のXML宣言とエンコーディングを保持
        $dom = new \DOMDocument('1.0', 'UTF-8');
        $dom->preserveWhiteSpace = true;
        $dom->formatOutput = false;
        $dom->encoding = 'UTF-8';
        
        // XMLをロード（エンティティとCDATAを処理）
        if (!@$dom->loadXML($xmlContent)) {
            $zip->close();
            throw new \Exception('Word文書のXML解析に失敗しました');
        }
        
        $xpath = new \DOMXPath($dom);
        $xpath->registerNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main');
        
        // w:t要素のみを取得して翻訳を適用
        $textNodes = $xpath->query('//w:t');
        foreach ($textNodes as $textNode) {
            $value = $textNode->nodeValue;
            if (isset($translations[$value]) && trim($value) !== '') {
                // 翻訳結果を取得
                $translatedValue = $translations[$value];
                
                // テキストノードを完全に置き換え（XMLエスケープは自動処理される）
                while ($textNode->hasChildNodes()) {
                    $textNode->removeChild($textNode->firstChild);
                }
                $textNode->appendChild($dom->createTextNode($translatedValue));
            }
        }
        
        // XMLを保存（宣言を含む）
        $newXmlContent = $dom->saveXML($dom->documentElement);
        // XML宣言を追加
        $newXmlContent = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n" . $newXmlContent;
        
        // ZIPを更新
        $zip->deleteName('word/document.xml');
        if (!$zip->addFromString('word/document.xml', $newXmlContent)) {
            $zip->close();
            throw new \Exception('翻訳後のXMLの追加に失敗しました');
        }
        
        $zip->close();

        return $outputPath;
    }

    /**
     * カスタムオプション設定フォーム
     */
    public function setCustomOptionForm(&$form)
    {
        $form->password('gemini_api_key', 'Gemini APIキー')
            ->required()
            ->help('Google AI StudioでGemini APIキーを取得してください: https://aistudio.google.com/app/apikey');
    }

    /**
     * インデックスページのHTMLを生成
     */
    private function generateIndexHtml($hasApiKey, $uploadUrl)
    {
        $csrfToken = csrf_token();
        $warningHtml = !$hasApiKey ? '<div class="alert alert-warning"><strong>注意:</strong> Gemini APIキーが設定されていません。プラグイン設定画面でAPIキーを設定してください。</div>' : '';

        return <<<HTML
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="csrf-token" content="{$csrfToken}">
    <style>
        .translator-container { max-width: 800px; margin: 0 auto; padding: 20px; }
        .upload-area { border: 2px dashed #ccc; border-radius: 8px; padding: 40px; text-align: center; background-color: #f9f9f9; cursor: pointer; transition: all 0.3s; }
        .upload-area:hover { border-color: #3c8dbc; background-color: #f0f8ff; }
        .upload-area.dragover { border-color: #3c8dbc; background-color: #e6f2ff; }
        .file-info { display: none; background: #e8f4f8; padding: 15px; border-radius: 8px; margin-top: 20px; }
        .config-section { display: none; margin-top: 20px; }
        .form-group { margin-bottom: 20px; }
        .form-group label { display: block; margin-bottom: 8px; font-weight: bold; color: #333; }
        .form-control { width: 100%; padding: 10px; border: 1px solid #ddd; border-radius: 4px; font-size: 14px; line-height: 1.6; box-sizing: border-box; }
        select.form-control { height: auto; min-height: 40px; }
        .btn { padding: 10px 24px; border: none; border-radius: 4px; cursor: pointer; font-size: 14px; transition: all 0.3s; }
        .btn-primary { background-color: #3c8dbc; color: white; }
        .btn-primary:hover { background-color: #357ca5; }
        .btn-primary:disabled { background-color: #ccc; cursor: not-allowed; }
        .progress-container { display: none; margin-top: 20px; }
        .progress-bar { width: 100%; height: 30px; background-color: #f0f0f0; border-radius: 15px; overflow: hidden; }
        .progress-fill { height: 100%; background: linear-gradient(90deg, #3c8dbc, #5cb85c); transition: width 0.3s; display: flex; align-items: center; justify-content: center; color: white; font-weight: bold; }
        .alert { padding: 12px 20px; border-radius: 4px; margin-top: 15px; }
        .alert-danger { background-color: #f2dede; color: #a94442; border: 1px solid #ebccd1; }
        .alert-success { background-color: #dff0d8; color: #3c763d; border: 1px solid #d6e9c6; }
        .alert-warning { background-color: #fcf8e3; color: #8a6d3b; border: 1px solid #faebcc; }
        .icon { font-size: 48px; color: #3c8dbc; margin-bottom: 10px; }
        .spinner { display: inline-block; width: 20px; height: 20px; border: 3px solid rgba(255,255,255,.3); border-radius: 50%; border-top-color: #fff; animation: spin 1s ease-in-out infinite; }
        @keyframes spin { to { transform: rotate(360deg); } }
    </style>
</head>
<body>
    <div class="translator-container">
        {$warningHtml}

        <div class="upload-area" id="uploadArea">
            <div class="icon">📄</div>
            <h3>ファイルを選択またはドラッグ&ドロップ</h3>
            <p>対応形式: Excel (.xlsx), Word (.docx)</p>
            <input type="file" id="fileInput" accept=".xlsx,.docx" style="display: none;">
            <button type="button" class="btn btn-primary" onclick="document.getElementById('fileInput').click()">ファイルを選択</button>
        </div>

        <div class="file-info" id="fileInfo">
            <h4>📎 <span id="fileName"></span></h4>
            <p>翻訳対象のユニークなテキスト: <strong id="uniqueCount">0</strong> 件</p>
        </div>

        <div class="config-section" id="configSection">
            <div class="form-group">
                <label for="targetLanguage">翻訳先の言語</label>
                <select class="form-control" id="targetLanguage">
                    <option value="日本語">日本語</option>
                    <option value="英語">英語</option>
                    <option value="中国語（簡体字）">中国語（簡体字）</option>
                    <option value="中国語（繁体字）">中国語（繁体字）</option>
                    <option value="韓国語">韓国語</option>
                    <option value="フランス語">フランス語</option>
                    <option value="ドイツ語">ドイツ語</option>
                    <option value="スペイン語">スペイン語</option>
                    <option value="イタリア語">イタリア語</option>
                    <option value="ポルトガル語">ポルトガル語</option>
                </select>
            </div>

            <div class="form-group">
                <label for="customInstruction">カスタム指示（オプション）</label>
                <textarea class="form-control" id="customInstruction" rows="3" placeholder="例：フォーマルな敬語を使ってください"></textarea>
                <small style="color: #666;">AIへの追加の指示を入力できます</small>
            </div>

            <button type="button" class="btn btn-primary" id="translateBtn" onclick="startTranslation()">
                <span id="translateBtnText">翻訳を開始</span>
                <span class="spinner" id="translateSpinner" style="display: none;"></span>
            </button>
        </div>

        <div class="progress-container" id="progressContainer">
            <h4>翻訳中...</h4>
            <div class="progress-bar">
                <div class="progress-fill" id="progressFill" style="width: 0%;"><span id="progressText">0%</span></div>
            </div>
            <p style="margin-top: 10px; color: #666;">お待ちください...</p>
        </div>

        <div id="alertContainer"></div>
    </div>

    <script>
        let uploadUrl = '{$uploadUrl}';
        let translateUrl = '';
        let downloadUrl = '';

        document.getElementById('fileInput').addEventListener('change', function(e) {
            if (e.target.files.length > 0) handleFileSelect(e.target.files[0]);
        });

        const uploadArea = document.getElementById('uploadArea');
        uploadArea.addEventListener('dragover', function(e) { e.preventDefault(); uploadArea.classList.add('dragover'); });
        uploadArea.addEventListener('dragleave', function() { uploadArea.classList.remove('dragover'); });
        uploadArea.addEventListener('drop', function(e) {
            e.preventDefault();
            uploadArea.classList.remove('dragover');
            if (e.dataTransfer.files.length > 0) handleFileSelect(e.dataTransfer.files[0]);
        });

        function handleFileSelect(file) {
            const formData = new FormData();
            formData.append('file', file);
            showAlert('info', 'ファイルを解析中...', false);

            fetch(uploadUrl, {
                method: 'POST',
                body: formData,
                headers: { 'X-CSRF-TOKEN': document.querySelector('meta[name="csrf-token"]').content }
            })
            .then(response => {
                console.log('Response status:', response.status);
                console.log('Response headers:', response.headers.get('content-type'));
                if (!response.ok) {
                    throw new Error('HTTP error ' + response.status);
                }
                const contentType = response.headers.get('content-type');
                if (!contentType || !contentType.includes('application/json')) {
                    return response.text().then(text => {
                        console.error('Non-JSON response:', text);
                        throw new Error('Server returned non-JSON response');
                    });
                }
                return response.json();
            })
            .then(data => {
                console.log('Upload response:', data);
                if (data.error) { 
                    const errorHtml = data.error.replace(/\\n/g, '<br>');
                    showAlert('danger', errorHtml); 
                    return; 
                }
                document.getElementById('fileName').textContent = file.name;
                document.getElementById('uniqueCount').textContent = data.uniqueCount;
                document.getElementById('fileInfo').style.display = 'block';
                document.getElementById('configSection').style.display = 'block';
                translateUrl = data.translateUrl;
                clearAlert();
            })
            .catch(error => {
                console.error('Upload error:', error);
                showAlert('danger', 'アップロード失敗: ' + error.message);
            });
        }

        function startTranslation() {
            const targetLanguage = document.getElementById('targetLanguage').value;
            const customInstruction = document.getElementById('customInstruction').value;
            const translateBtn = document.getElementById('translateBtn');
            
            translateBtn.disabled = true;
            document.getElementById('translateBtnText').textContent = '翻訳中...';
            document.getElementById('translateSpinner').style.display = 'inline-block';
            document.getElementById('progressContainer').style.display = 'block';

            let progress = 0;
            const progressInterval = setInterval(() => {
                progress += 1;
                if (progress <= 90) updateProgress(progress);
            }, 500);

            fetch(translateUrl, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'X-CSRF-TOKEN': document.querySelector('meta[name="csrf-token"]').content
                },
                body: JSON.stringify({ target_language: targetLanguage, custom_instruction: customInstruction })
            })
            .then(response => response.json())
            .then(data => {
                clearInterval(progressInterval);
                updateProgress(100);
                if (data.error) {
                    // エラーメッセージを改行付きで表示
                    const errorHtml = data.error.replace(/\\n/g, '<br>');
                    showAlert('danger', errorHtml);
                    translateBtn.disabled = false;
                    document.getElementById('translateBtnText').textContent = '翻訳を開始';
                    document.getElementById('translateSpinner').style.display = 'none';
                    document.getElementById('progressContainer').style.display = 'none';
                    return;
                }
                downloadUrl = data.downloadUrl;
                setTimeout(() => {
                    document.getElementById('progressContainer').style.display = 'none';
                    showTranslationComplete();
                }, 500);
            })
            .catch(error => {
                clearInterval(progressInterval);
                showAlert('danger', '翻訳失敗: ' + error.message);
                translateBtn.disabled = false;
                document.getElementById('translateBtnText').textContent = '翻訳を開始';
                document.getElementById('translateSpinner').style.display = 'none';
            });
        }

        function showTranslationComplete() {
            document.getElementById('alertContainer').innerHTML = '<div class="alert alert-success"><h4>✅ 翻訳完了！</h4><p>ファイルは正常に翻訳されました。</p><button type="button" class="btn btn-primary" onclick="downloadFile()" style="margin-top: 10px;">ファイルをダウンロード</button><button type="button" class="btn" onclick="location.reload()" style="margin-left: 10px; background: #6c757d; color: white;">別のファイルを翻訳</button></div>';
        }

        function downloadFile() {
            // fetchでファイルをBlobとして取得してダウンロード
            fetch(downloadUrl, {
                method: 'GET',
                headers: {
                    'X-CSRF-TOKEN': document.querySelector('meta[name="csrf-token"]').content
                }
            })
            .then(response => {
                if (!response.ok) {
                    throw new Error('ダウンロード失敗: ' + response.status);
                }
                return response.blob();
            })
            .then(blob => {
                // Blobからダウンロードリンクを作成
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                // URLからファイル名を抽出
                const filename = downloadUrl.split('/').pop();
                a.download = decodeURIComponent(filename);
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                document.body.removeChild(a);
            })
            .catch(error => {
                console.error('Download error:', error);
                showAlert('danger', 'ダウンロードに失敗しました: ' + error.message);
            });
        }

        function updateProgress(percent) {
            document.getElementById('progressFill').style.width = percent + '%';
            document.getElementById('progressText').textContent = percent + '%';
        }

        function showAlert(type, message, autoClear = true) {
            document.getElementById('alertContainer').innerHTML = '<div class="alert alert-' + type + '">' + message + '</div>';
            if (autoClear) setTimeout(clearAlert, 5000);
        }

        function clearAlert() {
            document.getElementById('alertContainer').innerHTML = '';
        }
    </script>
</body>
</html>
HTML;
    }
}
