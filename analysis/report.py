import os
import google.generativeai as genai
from docx import Document
import pandas as pd
from datetime import datetime
from pathlib import Path
import html
from itertools import combinations
import numpy as np
import asyncio
import logging
import re
import json
import sys

# ロギングの設定
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

class DocumentAnalyzer:
    def __init__(self, api_key, input_folder):
        """初期化"""
        self.setup_genai(api_key)
        self.input_folder = input_folder
        self.summaries = []
        self.comparisons = []
        self.discussion_points = []
        self.model = genai.GenerativeModel('gemini-pro')
        
        # 正規表現パターンを定義
        self.table_pattern = r'\|\s*(.+?)\s*\|[\r\n]+\|[-\s|]+\|[\r\n]+((?:\|.+\|[\r\n]+)+)'
        self.list_pattern = r'^\s*[-•]\s+(.+?)$'
        self.heading_pattern = r'^(\d+\.|#)\s*(.+?)$'
        self.emphasis_pattern = r'\*\*(.+?)\*\*'

    def setup_genai(self, api_key):
        """Gemini APIの設定"""
        try:
            genai.configure(api_key=api_key)
            logger.info("Gemini API setup completed successfully")
        except Exception as e:
            logger.error(f"API設定エラー: {e}")
            raise Exception(f"API設定エラー: {e}")

    def read_docx(self, file_path):
        """docxファイルを読み込んでテキストを抽出"""
        try:
            doc = Document(file_path)
            text = "\n".join([paragraph.text for paragraph in doc.paragraphs])
            logger.info(f"Successfully read file: {file_path}")
            return text
        except Exception as e:
            logger.error(f"ファイル読み込みエラー ({file_path}): {e}")
            return ""
            
            
    async def get_summary(self, text, file_name):
        """テキストの要約を生成"""
        try:
            prompt = f"""
            以下の文書を分析し、以下の形式で出力してください。
            出力形式:
            - 著者: [Sourceに記載された名称をの著者名として抽出。複数の場合は全て列挙]
            - 概要: [100文字程度の要約]
            - 主要な主張:
              1. [主張1]
              2. [主張2]
              3. [主張3]
            - 結論: [文書の結論や最終的な主張]
            
            文書:
            {text[:50000]}
            """
            
            response = self.model.generate_content(prompt)
            logger.info(f"Generated summary for: {file_name}")
            return {
                "file_name": file_name,
                "original_text": text,
                "summary": response.text
            }
        except Exception as e:
            logger.error(f"要約生成エラー ({file_name}): {e}")
            return None

    async def analyze_all_documents(self):
        """全文書の横断的な分析"""
        try:
            all_summaries = "\n\n".join([
                f"文書: {summary['file_name']}\n{summary['summary']}"
                for summary in self.summaries
            ])
            
            prompt = f"""
            以下の複数の文書を横断的に分析し、以下の形式で出力してください。
            
            出力形式:
            1. 全体概要表
            | 文書名 | 著者 | 主要な主張 | 結論 | キーワード |
            |--------|------|------------|------|------------|
            [各文書の情報を表形式で記載]

            2. 共通テーマと論点
            - [複数の文書で共通して議論されている主要なテーマを箇条書きで]
            
            3. 主張の比較表
            | テーマ | 各文書の立場・主張 |
            |--------|-------------------|
            [主要なテーマごとに各文書の立場を整理]

            4. 相違点の分析
            - [文書間の主要な意見の相違を箇条書きで]
            - [対立する意見がある場合はその内容]

            5. 共通認識
            - [全文書で共有されている前提や認識]
            - [一致している見解]

            6. 課題・今後の展望
            - [文書群から導かれる今後の課題]
            - [更なる検討が必要な点]

            文書群:
            {all_summaries}
            """
            
            response = self.model.generate_content(prompt)
            logger.info("Generated cross-document analysis")
            return response.text
        except Exception as e:
            logger.error(f"文書横断分析エラー: {e}")
            return None

    async def analyze_discussion_points(self):
        """議論ポイントの分析と優先度付け"""
        try:
            all_summaries = "\n\n".join([
                f"文書: {summary['file_name']}\n{summary['summary']}"
                for summary in self.summaries
            ])
            
            prompt = f"""
            以下の文書群から、重要な議論ポイントを抽出し、優先度付けして分析してください。
            
            出力形式:
            1. 最重要議論ポイント
            - [緊急性や重要性が特に高い議論ポイント]
              * 理由: [なぜ重要か]
              * 関連文書: [どの文書で言及されているか]

            2. 重要な議論ポイント
            - [重要度が高い議論ポイント]
              * 背景: [議論の背景]
              * 異なる見解: [存在する場合]

            3. 検討すべき議論ポイント
            - [長期的に検討が必要な議論ポイント]
              * 課題: [検討に必要な要素]

            4. 補足的な議論ポイント
            - [その他の関連する議論ポイント]

            文書群:
            {all_summaries}
            """
            
            response = self.model.generate_content(prompt)
            logger.info("Generated discussion points analysis")
            return response.text
        except Exception as e:
            logger.error(f"議論ポイント分析エラー: {e}")
            return None
            
    def format_content(self, text):
        """テキストコンテンツをマークダウンからHTMLに変換"""
        if not text:
            return ""
        
        try:
            # 改行を一時的に置き換え
            text = text.replace('\r\n', '\n')
            
            # セクションの処理
            section_pattern = r'^(#+)\s+(.+?)$'
            for match in re.finditer(section_pattern, text, re.MULTILINE):
                level = min(len(match.group(1)), 6)  # h1からh6まで
                heading = match.group(2)
                text = text.replace(
                    match.group(0), 
                    f'<h{level} class="content-heading">{heading}</h{level}>'
                )

            # 番号付きリストの処理
            numbered_list_pattern = r'(?:^\d+\.\s+(.+?)$\n?)+'
            text = re.sub(
                numbered_list_pattern,
                lambda m: '<ol>\n' + '\n'.join(
                    f'<li>{line.strip()}</li>' 
                    for line in re.findall(r'^\d+\.\s+(.+?)$', m.group(0), re.MULTILINE)
                ) + '\n</ol>',
                text,
                flags=re.MULTILINE
            )

            # 箇条書きの処理
            bullet_list_pattern = r'(?:^\s*[-*•]\s+(.+?)$\n?)+'
            text = re.sub(
                bullet_list_pattern,
                lambda m: '<ul>\n' + '\n'.join(
                    f'<li>{line.strip()}</li>'
                    for line in re.findall(r'^\s*[-*•]\s+(.+?)$', m.group(0), re.MULTILINE)
                ) + '\n</ul>',
                text,
                flags=re.MULTILINE
            )

            # 強調とイタリックの処理
            text = re.sub(r'\*\*\*(.+?)\*\*\*', r'<strong><em>\1</em></strong>', text)
            text = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', text)
            text = re.sub(r'\*(.+?)\*', r'<em>\1</em>', text)
            text = re.sub(r'___(.+?)___', r'<strong><em>\1</em></strong>', text)
            text = re.sub(r'__(.+?)__', r'<strong>\1</strong>', text)
            text = re.sub(r'_(.+?)_', r'<em>\1</em>', text)

            # コードブロックの処理
            text = re.sub(
                r'```(\w+)?\n(.*?)\n```',
                lambda m: f'<pre><code class="language-{m.group(1) or "plaintext"}">{html.escape(m.group(2))}</code></pre>',
                text,
                flags=re.DOTALL
            )
            text = re.sub(r'`([^`]+)`', r'<code>\1</code>', text)

            # 引用の処理
            text = re.sub(
                r'(?:^>\s+(.+?)$\n?)+',
                lambda m: '<blockquote>\n' + 
                         '\n'.join(re.findall(r'^>\s+(.+?)$', m.group(0), re.MULTILINE)) +
                         '\n</blockquote>',
                text,
                flags=re.MULTILINE
            )

            # 水平線の処理
            text = re.sub(r'^(?:-{3,}|\*{3,}|_{3,})$', '<hr>', text, flags=re.MULTILINE)

            # リンクの処理
            text = re.sub(r'\[([^\]]+)\]\(([^)]+)\)', r'<a href="\2" target="_blank">\1</a>', text)

            # 画像の処理
            text = re.sub(r'!\[([^\]]*)\]\(([^)]+)\)', r'<img src="\2" alt="\1">', text)

            # 段落の処理
            paragraphs = []
            current_paragraph = []
            
            for line in text.split('\n'):
                line = line.strip()
                if line:
                    current_paragraph.append(line)
                else:
                    if current_paragraph:
                        paragraph_text = ' '.join(current_paragraph)
                        if not any(paragraph_text.startswith(tag) for tag in 
                                 ['<h', '<ul', '<ol', '<blockquote', '<pre', '<hr']):
                            paragraphs.append(f'<p>{paragraph_text}</p>')
                        else:
                            paragraphs.append(paragraph_text)
                        current_paragraph = []
                    paragraphs.append('')
            
            # 最後の段落の処理
            if current_paragraph:
                paragraph_text = ' '.join(current_paragraph)
                if not any(paragraph_text.startswith(tag) for tag in 
                         ['<h', '<ul', '<ol', '<blockquote', '<pre', '<hr']):
                    paragraphs.append(f'<p>{paragraph_text}</p>')
                else:
                    paragraphs.append(paragraph_text)

            # 段落を結合
            text = '\n'.join(filter(None, paragraphs))

            return text

        except Exception as e:
            logger.error(f"コンテンツ整形中にエラーが発生: {e}")
            return html.escape(text)  # エラーが発生した場合は、エスケープした原文を返す
    
    def convert_markdown_table_to_html(self, text):
        """Markdownテーブルを整形されたHTMLテーブルに変換"""
        if not text:
            return text

        try:
            # テーブル全体を検出する正規表現
            table_pattern = r'(\|[^\n]+\|\n\|[-:\| ]+\|\n(?:\|[^\n]+\|\n?)+)'
            
            def convert_single_table(table_text):
                # テーブルの行を分割
                lines = table_text.strip().split('\n')
                if len(lines) < 3:  # ヘッダー行、区切り行、データ行が最低必要
                    return table_text

                try:
                    # ヘッダー行の処理
                    headers = [cell.strip() for cell in lines[0].split('|')[1:-1]]
                    
                    # 区切り行から位置揃えを取得
                    alignments = []
                    align_row = lines[1].strip()
                    align_cells = [cell.strip() for cell in align_row.split('|')[1:-1]]
                    
                    for cell in align_cells:
                        cell = cell.strip(':- ')
                        if cell.startswith(':') and cell.endswith(':'):
                            alignments.append('center')
                        elif cell.endswith(':'):
                            alignments.append('right')
                        else:
                            alignments.append('left')

                    # HTMLテーブルの構築
                    html_table = '<div class="table-container">\n'
                    html_table += '<table class="analysis-table">\n'

                    # ヘッダーの追加
                    html_table += '<thead>\n<tr>\n'
                    for i, header in enumerate(headers):
                        align = alignments[i] if i < len(alignments) else 'left'
                        html_table += f'<th class="text-{align}">{html.escape(header.strip())}</th>\n'
                    html_table += '</tr>\n</thead>\n'

                    # ボディ行の追加
                    html_table += '<tbody>\n'
                    for line in lines[2:]:  # ヘッダーと区切り行をスキップ
                        # 空行をスキップ
                        if not line.strip() or line.strip() == '|':
                            continue
                            
                        cells = [cell.strip() for cell in line.split('|')[1:-1]]
                        if not any(cells):  # 空の行をスキップ
                            continue
                            
                        html_table += '<tr>\n'
                        for i, cell in enumerate(cells):
                            align = alignments[i] if i < len(alignments) else 'left'
                            html_table += f'<td class="text-{align}">{html.escape(cell.strip())}</td>\n'
                        html_table += '</tr>\n'
                    
                    html_table += '</tbody>\n'
                    html_table += '</table>\n'
                    html_table += '</div>\n'
                    
                    return html_table
                except Exception as e:
                    logger.error(f"テーブル変換中にエラー: {e}")
                    return table_text

            # 検出されたすべてのテーブルを変換
            tables = re.finditer(table_pattern, text, re.MULTILINE)
            for match in tables:
                original_table = match.group(0)
                html_table = convert_single_table(original_table)
                text = text.replace(original_table, html_table)

            return text
        except Exception as e:
            logger.error(f"Markdownテーブル変換中にエラー: {e}")
            return text
    
    
    def generate_html_report(self):
        """HTMLレポートの生成"""
        try:
            current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            def process_content(content):
                if not content:
                    return ""
                # マークダウンテーブルをHTMLに変換し、その後フォーマットを適用
                processed = self.convert_markdown_table_to_html(content)
                return self.format_content(processed)

            html_content = f"""
            <!DOCTYPE html>
            <html lang="ja">
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>文書分析レポート</title>
                <style>
                    :root {{
                        --primary-color: #2c3e50;
                        --secondary-color: #3498db;
                        --background-color: #f5f7fa;
                        --card-background: #ffffff;
                        --text-color: #333333;
                        --border-color: #e1e4e8;
                    }}

                    body {{
                        font-family: 'Helvetica Neue', Arial, 'Hiragino Kaku Gothic ProN', 'Hiragino Sans', Meiryo, sans-serif;
                        line-height: 1.6;
                        max-width: 1200px;
                        margin: 0 auto;
                        padding: 20px;
                        background-color: var(--background-color);
                        color: var(--text-color);
                    }}

                    .card {{
                        background-color: var(--card-background);
                        border-radius: 8px;
                        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                        margin: 20px 0;
                        padding: 20px;
                    }}

                    h1, h2, h3, h4, h5, h6 {{
                        color: var(--primary-color);
                        border-bottom: 2px solid var(--border-color);
                        padding-bottom: 0.3em;
                        margin-top: 1.5em;
                        margin-bottom: 0.8em;
                    }}

                    .meta-info {{
                        color: #666;
                        font-size: 0.9em;
                        margin-bottom: 20px;
                    }}

                    .table-container {{
                        overflow-x: auto;
                        margin: 20px 0;
                        background: white;
                        border-radius: 8px;
                        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                        padding: 1px;
                    }}

                    .analysis-table {{
                        width: 100%;
                        border-collapse: collapse;
                        margin: 0;
                        background: white;
                        min-width: 600px;
                    }}

                    .analysis-table th {{
                        background-color: #f8f9fa;
                        color: var(--primary-color);
                        font-weight: bold;
                        padding: 12px 16px;
                        border: 1px solid var(--border-color);
                        white-space: nowrap;
                    }}

                    .analysis-table td {{
                        padding: 10px 16px;
                        border: 1px solid var(--border-color);
                        line-height: 1.5;
                    }}

                    .analysis-table tr:nth-child(even) {{
                        background-color: #f8f9fa;
                    }}

                    .text-left {{ text-align: left; }}
                    .text-center {{ text-align: center; }}
                    .text-right {{ text-align: right; }}

                    .nav-tabs {{
                        display: flex;
                        margin-bottom: 20px;
                        border-bottom: 2px solid var(--border-color);
                        background-color: var(--card-background);
                        border-radius: 8px 8px 0 0;
                        padding: 10px 10px 0 10px;
                    }}

                    .nav-tab {{
                        padding: 12px 24px;
                        cursor: pointer;
                        margin-right: 5px;
                        border: 1px solid var(--border-color);
                        border-bottom: none;
                        border-radius: 8px 8px 0 0;
                        background-color: #f8f9fa;
                        color: #666;
                    }}

                    .nav-tab:hover {{
                        background-color: #e9ecef;
                    }}

                    .nav-tab.active {{
                        background-color: var(--card-background);
                        border-bottom: 2px solid var(--card-background);
                        margin-bottom: -2px;
                        color: var(--primary-color);
                        font-weight: bold;
                    }}

                    .tab-content {{
                        display: none;
                        background-color: var(--card-background);
                        border-radius: 0 0 8px 8px;
                        padding: 20px;
                    }}

                    .tab-content.active {{
                        display: block;
                    }}

                    blockquote {{
                        border-left: 4px solid var(--secondary-color);
                        margin: 1em 0;
                        padding: 0.5em 1em;
                        background-color: #f8f9fa;
                    }}

                    code {{
                        background-color: #f8f9fa;
                        padding: 0.2em 0.4em;
                        border-radius: 3px;
                        font-family: monospace;
                    }}

                    ul, ol {{
                        padding-left: 2em;
                        margin: 1em 0;
                    }}

                    li {{
                        margin: 0.5em 0;
                    }}

                    .priority-high {{
                        border-left: 4px solid #e74c3c;
                    }}

                    .priority-medium {{
                        border-left: 4px solid #f39c12;
                    }}

                    .priority-low {{
                        border-left: 4px solid #2ecc71;
                    }}

                    @media (max-width: 768px) {{
                        body {{
                            padding: 10px;
                        }}
                        .card {{
                            padding: 15px;
                        }}
                        .nav-tab {{
                            padding: 8px 16px;
                        }}
                    }}
                </style>
                <script>
                    function showTab(tabId) {{
                        document.querySelectorAll('.tab-content').forEach(tab => {{
                            tab.classList.remove('active');
                        }});
                        document.querySelectorAll('.nav-tab').forEach(tab => {{
                            tab.classList.remove('active');
                        }});
                        document.getElementById(tabId).classList.add('active');
                        document.querySelector(`[onclick="showTab('${{tabId}}')"]`).classList.add('active');
                    }}
                </script>
            </head>
            <body>
                <h1>文書分析レポート</h1>
                <p class="meta-info">生成日時: {current_time}</p>
                
                <div class="card">
                    <h2>文書横断分析の概要</h2>
                    {process_content(self.cross_document_analysis) if hasattr(self, 'cross_document_analysis') else ''}
                </div>

                <div class="nav-tabs">
                    <div class="nav-tab active" onclick="showTab('tab-summary')">個別要約</div>
                    <div class="nav-tab" onclick="showTab('tab-themes')">テーマ別分析</div>
                    <div class="nav-tab" onclick="showTab('tab-discussion')">議論ポイント</div>
                </div>
            """

            # 個別要約タブ
            html_content += """
                <div id="tab-summary" class="tab-content active">
                    <h2>個別文書の要約</h2>
            """
            
            for summary in self.summaries:
                if summary:
                    html_content += f"""
                    <div class="card">
                        <h3>{html.escape(summary['file_name'])}</h3>
                        <div class="summary-content">
                            {process_content(summary['summary'])}
                        </div>
                    </div>
                    """

            # テーマ別分析タブ
            html_content += """
                </div>
                <div id="tab-themes" class="tab-content">
                    <h2>テーマ別分析</h2>
            """

            if hasattr(self, 'cross_document_analysis'):
                html_content += f"""
                <div class="card">
                    {process_content(self.cross_document_analysis)}
                </div>
                """

            # 議論ポイントタブ
            html_content += """
                </div>
                <div id="tab-discussion" class="tab-content">
                    <h2>議論ポイント</h2>
            """

            if self.discussion_points:
                html_content += f"""
                <div class="card">
                    {process_content(self.discussion_points)}
                </div>
                """

            html_content += """
                </div>
            </body>
            </html>
            """

            return html_content
        
        except Exception as e:
            logger.error(f"HTMLレポート生成中にエラーが発生: {e}")
            raise        

    def _generate_summaries_content(self):
        """個別要約タブのコンテンツを生成"""
        content = ""
        for summary in self.summaries:
            if summary:
                content += f"""
                <div class="card">
                    <h3>{html.escape(summary['file_name'])}</h3>
                    <div class="summary-content">
                        {self.format_content(html.escape(summary['summary']))}
                    </div>
                </div>
                """
        return content

    def _generate_themes_content(self):
        """テーマ別分析タブのコンテンツを生成"""
        if hasattr(self, 'cross_document_analysis'):
            return f"""
            <div class="card">
                {self.format_content(html.escape(self.cross_document_analysis))}
            </div>
            """
        return ""

    def _generate_discussion_content(self):
        """議論ポイントタブのコンテンツを生成"""
        if self.discussion_points:
            return f"""
            <div class="card">
                {self.format_content(html.escape(self.discussion_points))}
            </div>
            """
        return ""
        
    async def process_documents(self):
        """文書の処理メインフロー"""
        try:
            # docxファイルを検索
            docx_files = [f for f in os.listdir(self.input_folder) if f.endswith('.docx')]
            logger.info(f"見つかったdocxファイル数: {len(docx_files)}")

            if not docx_files:
                logger.warning("処理対象のdocxファイルが見つかりませんでした")
                return

            # 進捗表示用の総ファイル数
            total_files = len(docx_files)
            
            # 各文書の要約を生成
            for index, file_name in enumerate(docx_files, 1):
                logger.info(f"処理中 ({index}/{total_files}): {file_name}")
                file_path = os.path.join(self.input_folder, file_name)
                
                # ファイル読み込み
                text = self.read_docx(file_path)
                if not text:
                    logger.warning(f"ファイル {file_name} は空か読み込めませんでした")
                    continue

                # 要約生成
                try:
                    summary = await self.get_summary(text, file_name)
                    if summary:
                        self.summaries.append(summary)
                        logger.info(f"要約生成完了: {file_name}")
                    else:
                        logger.warning(f"要約生成失敗: {file_name}")
                except Exception as e:
                    logger.error(f"要約生成中にエラーが発生 ({file_name}): {e}")
                    continue

            if not self.summaries:
                logger.error("有効な要約が生成されませんでした")
                return

            # 中間結果のバックアップ
            await self.save_intermediate_results()

            # 文書横断分析を実行
            logger.info("文書横断分析を開始")
            try:
                self.cross_document_analysis = await self.analyze_all_documents()
                if not self.cross_document_analysis:
                    logger.warning("文書横断分析の結果が空でした")
            except Exception as e:
                logger.error(f"文書横断分析中にエラーが発生: {e}")

            # 議論ポイントを分析
            logger.info("議論ポイント分析を開始")
            try:
                self.discussion_points = await self.analyze_discussion_points()
                if not self.discussion_points:
                    logger.warning("議論ポイント分析の結果が空でした")
            except Exception as e:
                logger.error(f"議論ポイント分析中にエラーが発生: {e}")

            # HTMLレポートを生成して保存
            logger.info("HTMLレポート生成を開始")
            try:
                html_report = self.generate_html_report()
                output_path = os.path.join(self.input_folder, 'document_analysis_report.html')
                
                with open(output_path, 'w', encoding='utf-8') as f:
                    f.write(html_report)
                logger.info(f"レポートを保存しました: {output_path}")
                
                # 処理完了サマリーを表示
                print("\n=== 処理完了 ===")
                print(f"処理したファイル数: {len(docx_files)}")
                print(f"生成した要約数: {len(self.summaries)}")
                print(f"出力ファイル: {output_path}")
                
            except Exception as e:
                logger.error(f"レポート生成・保存中にエラーが発生: {e}")
                raise

        except Exception as e:
            logger.error(f"処理全体でエラーが発生: {e}")
            raise

    async def save_intermediate_results(self):
        """中間結果の保存（エラー発生時のバックアップ用）"""
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_dir = os.path.join(self.input_folder, 'backups')
            os.makedirs(backup_dir, exist_ok=True)
            
            backup_data = {
                'summaries': self.summaries,
                'cross_document_analysis': self.cross_document_analysis if hasattr(self, 'cross_document_analysis') else None,
                'discussion_points': self.discussion_points if hasattr(self, 'discussion_points') else None
            }
            
            backup_path = os.path.join(backup_dir, f'analysis_backup_{timestamp}.json')
            with open(backup_path, 'w', encoding='utf-8') as f:
                json.dump(backup_data, f, ensure_ascii=False, indent=2)
            
            logger.info(f"中間結果を保存しました: {backup_path}")
            
        except Exception as e:
            logger.error(f"中間結果の保存中にエラーが発生: {e}")

# メイン実行部分
async def main():
    """メイン実行関数"""
    try:
        # コマンドライン引数の処理
        parser = argparse.ArgumentParser(description='文書分析ツール')
        parser.add_argument('--api-key', help='Google AI Studio APIキー')
        parser.add_argument('--input-folder', help='分析対象の文書があるフォルダパス')
        parser.add_argument('--debug', action='store_true', help='デバッグモードを有効にする')
        args = parser.parse_args()

        # ロギングレベルの設定
        if args.debug:
            logging.getLogger().setLevel(logging.DEBUG)

        # 設定値の取得
        api_key = args.api_key or os.getenv('GOOGLE_AI_API_KEY') or 'API KEY'
        input_folder = args.input_folder or 'folder path'

        # 入力値の検証
        if api_key == 'your-api-key':
            raise ValueError("APIキーが設定されていません。--api-keyオプションまたは環境変数GOOGLE_AI_API_KEYで指定してください。")
        if not os.path.exists(input_folder):
            raise ValueError(f"指定されたフォルダが存在しません: {input_folder}")

        print("\n=== 文書分析開始 ===")
        print(f"対象フォルダ: {input_folder}")
        
        # 処理開始時刻を記録
        start_time = datetime.now()

        # DocumentAnalyzerのインスタンス作成と実行
        analyzer = DocumentAnalyzer(api_key, input_folder)
        await analyzer.process_documents()

        # 処理時間の計算と表示
        end_time = datetime.now()
        processing_time = end_time - start_time
        print(f"\n処理時間: {processing_time}")

    except ValueError as ve:
        logger.error(f"設定エラー: {ve}")
        print(f"\nエラー: {ve}")
        print("正しい設定値を指定してください。")
        sys.exit(1)
    except Exception as e:
        logger.error(f"予期せぬエラーが発生: {e}")
        print(f"\n予期せぬエラーが発生しました: {e}")
        print("詳細はログファイルを確認してください。")
        sys.exit(1)

if __name__ == "__main__":
    import argparse
    asyncio.run(main())
