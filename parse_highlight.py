#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
採録論文ハイライトPDFをパースしてExcelに出力するスクリプト

使用方法:
    python parse_highlight.py <PDFファイル> [出力Excelファイル]

依存ライブラリのインストール:
    pip install pymupdf openpyxl requests beautifulsoup4
"""

import re
import sys
import argparse
import unicodedata
from pathlib import Path

try:
    import fitz  # PyMuPDF
except ImportError:
    print("Error: PyMuPDF がインストールされていません")
    print("インストール: pip install pymupdf")
    sys.exit(1)

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side
except ImportError:
    print("Error: openpyxl がインストールされていません")
    print("インストール: pip install openpyxl")
    sys.exit(1)

try:
    import requests
    from bs4 import BeautifulSoup
    HAS_WEB_LIBS = True
except ImportError:
    HAS_WEB_LIBS = False
    print("Warning: requests/beautifulsoup4 がインストールされていません")
    print("DOI自動取得を使用するには: pip install requests beautifulsoup4")
    print("開始ページ番号からDOIを推測します")
    print()


def normalize_title(title: str) -> str:
    """タイトルを正規化して比較しやすくする"""
    # Unicode正規化
    title = unicodedata.normalize('NFKC', title)
    # 空白・改行を除去
    title = re.sub(r'\s+', '', title)
    # 全角ダッシュを統一
    title = title.replace('—', '-').replace('–', '-').replace('―', '-').replace('̶', '-')
    # 小文字化（英字部分）
    title = title.lower()
    return title


def fetch_doi_from_jstage(vol: int, no: int) -> dict[str, str]:
    """
    J-STAGEから論文タイトルとDOIの対応を取得する

    Args:
        vol: 巻番号
        no: 号番号

    Returns:
        {正規化タイトル: DOI} の辞書
    """
    if not HAS_WEB_LIBS:
        return {}

    url = f"https://www.jstage.jst.go.jp/browse/jsise/{vol}/{no}/_contents/-char/ja"
    print(f"J-STAGEからDOI情報を取得中: {url}")

    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36'
        }
        response = requests.get(url, headers=headers, timeout=30)
        response.raise_for_status()

        soup = BeautifulSoup(response.text, 'html.parser')

        # 論文エントリを探す
        doi_map = {}

        # 各論文のリンクを探す
        # J-STAGEの構造: div.searchlist-title > a にタイトル
        # DOIは data-doi 属性または別の要素に含まれる
        articles = soup.find_all('div', class_='searchlist-title')

        for article in articles:
            title_link = article.find('a')
            if title_link:
                title = title_link.get_text(strip=True)
                normalized = normalize_title(title)

                # DOIを探す - 親要素から探す
                parent = article.find_parent('div', class_='searchlist-col-item')
                if parent:
                    doi_elem = parent.find('a', href=re.compile(r'doi\.org'))
                    if doi_elem:
                        doi_url = doi_elem.get('href', '')
                        # DOIをフルURL形式で抽出（https://doi.org/10.14926/jsise.43.26 形式）
                        doi_match = re.search(r'(https?://doi\.org/10\.\d+/[^\s]+)', doi_url)
                        if doi_match:
                            doi_map[normalized] = doi_match.group(1)
                        else:
                            # DOI識別子のみの場合はURLに変換
                            doi_match = re.search(r'(10\.\d+/[^\s]+)', doi_url)
                            if doi_match:
                                doi_map[normalized] = f"https://doi.org/{doi_match.group(1)}"

        # 別のパターンでも探す
        if not doi_map:
            # すべてのDOIリンクを探す
            doi_links = soup.find_all('a', href=re.compile(r'doi\.org/10\.14926/jsise'))
            for link in doi_links:
                doi_url = link.get('href', '')
                # フルURL形式で抽出
                doi_match = re.search(r'(https?://doi\.org/10\.14926/jsise\.\d+\.\d+)', doi_url)
                if doi_match:
                    doi = doi_match.group(1)
                else:
                    # DOI識別子のみの場合はURLに変換
                    doi_match = re.search(r'(10\.14926/jsise\.\d+\.\d+)', doi_url)
                    if doi_match:
                        doi = f"https://doi.org/{doi_match.group(1)}"
                    else:
                        continue

                # 近くのタイトルを探す
                # 親要素を遡ってタイトルを探す
                container = link.find_parent(['div', 'li', 'article'])
                if container:
                    title_elem = container.find(['h3', 'h4', 'a'], class_=re.compile(r'title'))
                    if not title_elem:
                        title_elem = container.find('a')
                    if title_elem:
                        title = title_elem.get_text(strip=True)
                        if len(title) > 10:  # 短すぎるタイトルは除外
                            normalized = normalize_title(title)
                            doi_map[normalized] = doi

        print(f"  取得したDOI数: {len(doi_map)}")
        return doi_map

    except requests.RequestException as e:
        print(f"  Warning: J-STAGEからの取得に失敗しました: {e}")
        return {}


def guess_doi_from_page(vol: int, start_page: int) -> str:
    """
    巻番号と開始ページ番号からDOIのURLを推測する

    教育システム情報学会誌のDOIパターン: https://doi.org/10.14926/jsise.{巻}.{開始ページ}
    """
    return f"https://doi.org/10.14926/jsise.{vol}.{start_page}"


def extract_vol_no_from_pdf(text: str) -> tuple[int, int] | None:
    """
    PDFテキストから巻号を抽出する

    例: 「Vol. 43, No. 1」から (43, 1) を返す
    """
    match = re.search(r'Vol\.\s*(\d+),?\s*No\.\s*(\d+)', text)
    if match:
        return int(match.group(1)), int(match.group(2))
    return None


def extract_text_from_pdf(pdf_path: str) -> str:
    """PDFからテキストを抽出する（PyMuPDF使用）"""
    text_parts = []
    doc = fitz.open(pdf_path)
    for page in doc:
        text = page.get_text()
        if text:
            # ページ末尾の余計な空白行を除去
            text = text.rstrip()
            text_parts.append(text)
    doc.close()
    # ページ間は空白なしで結合（改行1つのみ）
    return "\n".join(text_parts)


def parse_papers(text: str) -> list[dict]:
    """
    テキストから論文情報をパースする
    """
    papers = []

    # 論文エントリのパターン（種別とページ範囲）
    # 【一般論文】（pp. 26‒38）や【ショートノート】（pp. 49‒54）など
    entry_pattern = r'【(一般論文|実践論文|ショートノート|実践速報)】（pp\.\s*(\d+)[‒\-–](\d+)）'

    # 全エントリの位置を特定
    entries = list(re.finditer(entry_pattern, text))

    for idx, match in enumerate(entries):
        paper = {
            '種別': match.group(1),
            '開始ページ': int(match.group(2)),
            '終了ページ': int(match.group(3)),
            'DOI': '',  # PDF内に記載なし
            'タイトル': '',
            '著者': '',
            'カテゴリ': '',
            '分野': '',
            '問い': '',
            '面白い点': ''
        }

        # このエントリのテキスト範囲を決定
        start_pos = match.end()
        if idx + 1 < len(entries):
            end_pos = entries[idx + 1].start()
        else:
            end_pos = len(text)

        entry_text = text[start_pos:end_pos]

        # タイトルと著者の抽出
        # PDFでは「タイトル\n著者名　姓名（所属）ほか」の形式
        # 姓名の間に全角スペースがある

        # 著者情報を探す：姓＋全角スペース＋名＋（所属）＋ほか
        # 例：島崎　俊介（電気通信大学）ほか
        # \u4e00-\u9fff: CJK統合漢字
        # \u3040-\u309f: ひらがな
        # \u30a0-\u30ff: カタカナ
        # \u3000: 全角スペース (IDEOGRAPHIC SPACE)
        # \uff08, \uff09: 全角括弧
        author_match = re.search(
            r'\n([\u4e00-\u9fff]{1,4})\u3000([\u4e00-\u9fff\u3040-\u309f\u30a0-\u30ff]{1,4})\uff08([^\uff09]+)\uff09(ほか)?',
            entry_text
        )

        if author_match:
            author_start = author_match.start()
            # 著者情報より前の部分がタイトル（先頭の改行を除去）
            title_text = entry_text[:author_start]
            # タイトルから不要な改行・空白を除去
            title = re.sub(r'[\s\n]+', '', title_text.strip())
            paper['タイトル'] = title

            # 著者情報を構築
            surname = author_match.group(1)
            given_name = author_match.group(2)
            affiliation = author_match.group(3)
            has_others = author_match.group(4)

            if has_others:
                paper['著者'] = f"{surname}　{given_name}（{affiliation}）ほか"
            else:
                paper['著者'] = f"{surname}　{given_name}（{affiliation}）"

        # カテゴリの抽出
        category_match = re.search(r'カテゴリ[：:]\s*(.+?)(?=\n|分\s*野)', entry_text)
        if category_match:
            paper['カテゴリ'] = category_match.group(1).strip()

        # 分野の抽出
        field_match = re.search(r'分\s*野[：:]\s*(.+?)(?=\n本論文|$)', entry_text, re.DOTALL)
        if field_match:
            field = field_match.group(1).strip()
            # 不要な改行を除去
            field = re.sub(r'\s+', '', field)
            # カンマの後にスペースを入れる
            field = re.sub(r'，', '，', field)
            paper['分野'] = field

        # 「問い」の抽出
        question_match = re.search(
            r'本論文の扱う「問い」\s*(.+?)(?=本論文のここが面白い！|$)',
            entry_text,
            re.DOTALL
        )
        if question_match:
            question = question_match.group(1).strip()
            # 改行を整理（RQ項目の改行は保持，それ以外は除去）
            lines = question.split('\n')
            cleaned_questions = []
            current_item = ''

            for line in lines:
                line = line.strip()
                if not line:
                    continue
                # ページ番号やヘッダーをスキップ
                if re.match(r'^\d+$', line):
                    continue
                if '教育システム情報学会誌' in line:
                    continue
                if line == '採録論文ハイライト':
                    continue

                # RQ項目の開始（・で始まる行）
                if re.match(r'^・', line):
                    if current_item:
                        cleaned_questions.append(current_item)
                    current_item = line
                else:
                    # 継続行（改行を除去して結合）
                    current_item += line

            if current_item:
                cleaned_questions.append(current_item)

            # RQ形式がない場合は、全体を1つのテキストとして結合
            if cleaned_questions:
                paper['問い'] = '\n'.join(cleaned_questions)
            else:
                # RQ形式でない場合、改行を除去して1つのテキストに
                cleaned = ''.join(line.strip() for line in question.split('\n')
                                  if line.strip() and not re.match(r'^\d+$', line.strip())
                                  and '教育システム情報学会誌' not in line
                                  and line.strip() != '採録論文ハイライト')
                paper['問い'] = cleaned

        # 「面白い！」の抽出
        interesting_match = re.search(
            r'本論文のここが面白い！\s*(.+?)(?=【|$)',
            entry_text,
            re.DOTALL
        )
        if interesting_match:
            interesting = interesting_match.group(1).strip()
            # 不要な改行を除去
            lines = interesting.split('\n')
            valid_lines = []

            for line in lines:
                line = line.strip()
                if not line:
                    continue
                # ページ番号やヘッダー/フッターをスキップ
                if re.match(r'^\d+$', line):
                    continue
                if '教育システム情報学会誌' in line:
                    continue
                if line == '採録論文ハイライト':
                    continue
                if re.match(r'^pp\.\s*\d+', line):
                    continue
                if re.match(r'^Vol\.\s*\d+', line):
                    continue

                valid_lines.append(line)

            # 全ての行を結合（PDFの改行は基本的に不要）
            paper['面白い点'] = ''.join(valid_lines)

        papers.append(paper)

    return papers


def save_to_excel(papers: list[dict], output_path: str):
    """論文情報をExcelファイルに保存する"""
    wb = Workbook()
    ws = wb.active
    ws.title = "採録論文ハイライト"

    # ヘッダー
    headers = ['種別', '開始ページ', '終了ページ', 'DOI', 'タイトル', '著者',
               'カテゴリ', '分野', '問い', '面白い点']

    # ヘッダースタイル
    header_font = Font(bold=True)
    header_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # ヘッダー行を書き込み
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.alignment = header_alignment
        cell.border = thin_border

    # データ行を書き込み
    cell_alignment = Alignment(vertical='top', wrap_text=True)

    for row_idx, paper in enumerate(papers, 2):
        for col_idx, header in enumerate(headers, 1):
            value = paper.get(header, '')
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.alignment = cell_alignment
            cell.border = thin_border

    # 列幅を調整
    column_widths = {
        'A': 12,   # 種別
        'B': 10,   # 開始ページ
        'C': 10,   # 終了ページ
        'D': 15,   # DOI
        'E': 40,   # タイトル
        'F': 25,   # 著者
        'G': 20,   # カテゴリ
        'H': 30,   # 分野
        'I': 50,   # 問い
        'J': 80,   # 面白い点
    }

    for col_letter, width in column_widths.items():
        ws.column_dimensions[col_letter].width = width

    # 行の高さを自動調整（最小値設定）
    for row in range(2, len(papers) + 2):
        ws.row_dimensions[row].height = 100

    wb.save(output_path)
    print(f"Excelファイルを保存しました: {output_path}")


def main():
    parser = argparse.ArgumentParser(
        description='採録論文ハイライトPDFをパースしてExcelに出力する'
    )
    parser.add_argument('pdf_file', help='入力PDFファイル')
    parser.add_argument('output_file', nargs='?', default=None,
                        help='出力Excelファイル（省略時はPDFと同名の.xlsx）')

    args = parser.parse_args()

    pdf_path = Path(args.pdf_file)
    if not pdf_path.exists():
        print(f"Error: ファイルが見つかりません: {pdf_path}")
        sys.exit(1)

    if args.output_file:
        output_path = Path(args.output_file)
    else:
        output_path = pdf_path.with_suffix('.xlsx')

    print(f"PDFを読み込み中: {pdf_path}")

    # PDFからテキスト抽出
    raw_text = extract_text_from_pdf(str(pdf_path))

    # 巻号を抽出
    vol_no = extract_vol_no_from_pdf(raw_text)
    if vol_no:
        vol, no = vol_no
        print(f"検出した巻号: Vol.{vol} No.{no}")
    else:
        print("Warning: 巻号を検出できませんでした。DOIは開始ページから推測します。")
        vol, no = None, None

    # 論文情報をパース（改行が重要なため、clean_textは適用しない）
    papers = parse_papers(raw_text)

    print(f"検出した論文数: {len(papers)}")

    # J-STAGEからDOI情報を取得
    doi_map = {}
    if vol and no:
        doi_map = fetch_doi_from_jstage(vol, no)

    # DOIを設定
    for paper in papers:
        normalized_title = normalize_title(paper['タイトル'])

        # タイトルマッチングでDOIを探す
        doi_found = False
        for jstage_title, doi in doi_map.items():
            # 部分一致でも検索（副題が異なる場合など）
            if normalized_title in jstage_title or jstage_title in normalized_title:
                paper['DOI'] = doi
                doi_found = True
                break
            # より緩い一致（タイトルの最初の部分）
            if len(normalized_title) > 20 and len(jstage_title) > 20:
                if normalized_title[:20] == jstage_title[:20]:
                    paper['DOI'] = doi
                    doi_found = True
                    break

        # マッチしなかった場合は開始ページから推測
        if not doi_found and vol:
            paper['DOI'] = guess_doi_from_page(vol, paper['開始ページ'])
            print(f"  DOI推測: {paper['タイトル'][:25]}... -> {paper['DOI']}")

    for i, paper in enumerate(papers, 1):
        doi_status = "✓" if paper['DOI'] else "✗"
        print(f"  {i}. [{paper['種別']}] {paper['タイトル'][:30]}... [{doi_status}]")

    # Excelに保存
    save_to_excel(papers, str(output_path))


if __name__ == '__main__':
    main()
