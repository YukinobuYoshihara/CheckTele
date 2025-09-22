
import pandas as pd
from datetime import datetime, date
import os
import numpy as np
import re
import glob
from typing import Optional, Union, Dict
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell

# ==============================================================================
# 設定クラス
# ==============================================================================

class ReportConfig:
    """
    レポート作成に関するすべての設定値を一元管理するクラス。

    このクラスは、ファイルパス、シート名、読み込み設定、Excelのスタイル定義など、
    スクリプト全体で使用される定数や設定値を属性として一元的に管理します。

    Attributes:
        FILES (dict): 入力ファイルのパス情報。
        SHEET_NAMES (dict): 出力Excelのシート名定義。
        LOAD_CONFIG (dict): 各種データ読み込み時の設定値。
        COL_NAMES_CONSTRUCTION_PREP (list): 構築準備データの列名リスト。
        PARAM_PHASE_DEFINITIONS (dict): パラメータシート作成のフェーズ定義。
        CONSTRUCTION_PHASE_DEFINITIONS (dict): 構築準備のフェーズ定義。
        KUKAN_TO_PHASES_MAP (dict): 区分ごとのフェーズ定義マッピング。
        PROCESS_DEFINITIONS (dict): 進捗状況シート用のプロセス定義。
        DOCUMENT_ORDER (list): 設計書名の表示順リスト。
        SHEET_STYLES (dict): シートごとのスタイルと集計設定。
    """
    # --- ファイルパス ---
    FILES = {
        'online_progress': 'テレ為替基盤_内部スケジュール(ID)_(オンラインのみ)20250625a.xlsx',
        'other_progress': 'テレ為替基盤_内部スケジュール(ID)サマリ_20250702.xlsx',
        'issue_list': '詳細設計事前検討一覧_マージ版.xlsx',
        'construction_schedule': 'テレ為替基盤_構築準備スケジュール_20250908.xlsx'
    }

    # --- シート名定義 ---
    SHEET_NAMES = {
        "progress": "進捗状況",
        "issue": "課題状況",
        "new_transfer": "新ファイル転送進捗状況",
        "missing_page": "頁数未設定",
        "construction_prep": "構築準備"
    }

    # --- 読み込み設定 ---
    LOAD_CONFIG = {
        'START_ROW_ONLINE': 14, 'START_ROW': 54, 'START_COLUMN': 3, 'END_COLUMN': 14,
        'START_COLUMN_NEW_FILE_TRANSFER': 2, 'END_COLUMN_NEW_FILE_TRANSFER': 13,
        'START_ROW_ISSUE_LIST': 5, 'PROGRESS_SHEET_NAME': 'テレ為替',
        'NEW_FILE_TRANSFER_PROGRESS_SHEET_NAME': '新ファイル転送',
        'CONSTRUCTION_PREP_SHEET_NAME': 0,
        'START_ROW_CONSTRUCTION_PREP': 14,
        'USE_COLS_CONSTRUCTION_PREP': 'D:O'
    }

    COL_NAMES_CONSTRUCTION_PREP = [
        '区分','対象', '作業区分', '作業項目','開始予定日','終了予定日','開始実績日',
        '終了実績日','予定頁数','実績頁数','担当','状況'
    ]

    # --- 区分ごとのフェーズ定義 ---
    PARAM_PHASE_DEFINITIONS = {
        '準備フェーズ': ['雛形/パラメータ決定', '情報収集', '雛形(PV)B.U.', '課題抽出/QA発出',  '課題抽出/QA発出（中断）', '課題抽出/QA発出(再開）', '非互換影響確認', '雛形(PV)認識合わせ', 'ヒアリング項目作成', '非互換確認', '環境準備', 'APヒアリング', '雛形入手or作成', '(雛形すり合わせ）', '非互換確認／修正', '雛形（集配信定義除く）B.U.', '雛形作成'],
        'パラメータシート作成フェーズ': ['パラメータ仮埋め', 'パラメータ仮埋め（コア）','パラメータシート修正', 'パラメタ見直し', '非互換/変更要件反映', '中間すり合わせ', 'パラメータ見直し', 'パラメータ仮決め', 'パラメータ仮決め（ジョブ除く）', 'パラメータ入力', '作成', 'パラメータ修正'],
        'T内Revフェーズ': ['T内Rev', 'T'],
        'GLRevフェーズ': ['GLRev', 'GL'],
        'デザインRevフェーズ': ['デザインRev'],
        'コア→モア、他センタ展開フェーズ': ['コア→モア、他センタ展開']
    }

    CONSTRUCTION_PHASE_DEFINITIONS = {
        '構築手順、定義体作成フェーズ': ['構築手順、定義体作成'],
        'T内Revフェーズ': ['T内Rev', 'T']
    }

    KUKAN_TO_PHASES_MAP = {
        "パラメータシート作成": PARAM_PHASE_DEFINITIONS,
        "構築準備": CONSTRUCTION_PHASE_DEFINITIONS
    }

    # --- 進捗状況シート用のプロセス定義 ---
    PROCESS_DEFINITIONS = {
        'writing': ['設計書執筆（骨子への下書き）', '設計書執筆（別紙作成下書き）','設計書修正'],
        'finalize': ['清書(PullRequest依頼まで)', '清書(PullRequest依頼⇒修正含む)', '設計書執筆（MarkDown清書PullRequest依頼まで）&確認結果修正'],
        'team_review': ['T内レビュー', 'チーム内レビュー'], 'gl_review': ['GLレビュー'], 'psl_review': ['PSLレビュー']
    }

    # --- ドキュメント順序 ---
    DOCUMENT_ORDER = [
        'オンライン処理方式詳細設計書','データベース詳細設計書','帳票処理方式詳細設計書','クラスタミドルウェア詳細設計書',
        'セキュリティ詳細設計書','システム運転管理詳細設計書','システム監視方式詳細設計書','ジョブネット設計規約',
        'バックアップ処理方式詳細設計書','リリース管理方式詳細設計書','ログ管理方式詳細設計書','処理実績管理方式詳細設計書',
        'システム運用様式','OS詳細設計書','HW設備詳細設計書','ストレージ詳細設計書','ネットワーク詳細設計書',
        '仮想化基盤詳細設計書','端末詳細設計書','バッチ処理方式詳細設計書','バックアップ方式詳細設計書',
        '新F転(オープンサーバ/AP基盤編)','新F転(AP基盤編)','新F転(オープンサーバ)','共通'
    ]

    # --- シートごとのスタイルと集計設定 ---
    SHEET_STYLES = {
        SHEET_NAMES["progress"]: {
            "column_widths": {'A:A': 34.3, 'B:F': 11.4, 'G:H': 14.4, 'I:O': 11.4, 'P:Q': 14.4, 'R:AC': 11.4},
            "sum_cols": [1, 2, 3, 5, 8, 9, 11, 12, 14, 17, 18, 20, 21, 22, 23, 24, 25, 26, 27, 28],
            "ratio_cols": [(3, 1, 4), (9, 8, 10), (12, 1, 13), (18, 17, 19)],
            "diff_cols": list(range(1, 6)) + list(range(8, 15)) + list(range(17, 29)),
            "percent_cols": [4, 10, 13, 19], "date_cols": [6, 15],
            "highlight_rows": {"column": "設計書名", "values": ["帳票処理方式詳細設計書", "システム運用様式", "端末詳細設計書"]}
        },
        SHEET_NAMES["issue"]: {
            "column_widths": {'A:A': 35, 'B:E': 15, 'F:H': 25}, "sum_cols": [1, 2, 3, 4], "ratio_cols": [], "diff_cols": [1, 2, 3, 4]
        },
        SHEET_NAMES["new_transfer"]: {
            "column_widths": {'A:A': 34.3, 'B:K': 11.4, 'L:T': 11.4}, "sum_cols": [1, 2, 3, 5, 8, 9, 11, 12, 13, 14, 15, 16, 17, 18],
            "ratio_cols": [(3, 1, 4), (9, 8, 10)], "diff_cols": list(range(1, 6)) + list(range(8, 11)) + list(range(11, 19)),
            "percent_cols": [4, 10], "date_cols": [6]
        },
        SHEET_NAMES["missing_page"]: {
            "column_widths": {'A:A': 30, 'B:B': 30, 'C:G': 15, 'H:I': 12, 'J:K': 15}, "date_cols": [3, 4, 5, 6]
        }
    }

# ==============================================================================
# データ抽出・加工・集計系の関数
# ==============================================================================
def extract_columns_by_range(dataframe: pd.DataFrame, start_col_index: int, end_col_index: int) -> pd.DataFrame:
    """
    指定した列範囲をDataFrameから抽出し、列名を標準化して返します。

    Args:
        dataframe (pd.DataFrame): 元データフレーム。
        start_col_index (int): 抽出開始列インデックス（0始まり）。
        end_col_index (int): 抽出終了列インデックス（0始まり、endは含まない）。

    Returns:
        pd.DataFrame: 抽出・整形済みのデータフレーム。

    Raises:
        IndexError: インデックスが範囲外または不正な場合。
    """
    num_cols = len(dataframe.columns)
    if not (0 <= start_col_index < num_cols and 0 <= end_col_index <= num_cols and start_col_index < end_col_index):
        raise IndexError(f"列のインデックスが範囲外または不正です。start:{start_col_index}, end:{end_col_index}, total:{num_cols}")
    extracted_df = dataframe.iloc[:, start_col_index:end_col_index]
    new_column_names = ['document_name','process_name','work_item','start_date_planned','end_date_planned','start_date_actual','end_date_actual','page_planned','page_actual','assignee','status']
    extracted_df.columns = new_column_names
    return extracted_df

def extract_columns_by_range_issue(dataframe: pd.DataFrame) -> pd.DataFrame:
    """
    課題管理用のDataFrameから不連続な範囲の列を抽出し、新しい列名を付けて返します。

    Args:
        dataframe (pd.DataFrame): 元データフレーム。

    Returns:
        pd.DataFrame: 抽出・整形済みの課題データフレーム。

    Raises:
        ValueError: 抽出された列数と新しい列名の数が一致しない場合。
    """
    target_indices = np.r_[0:5, 10:20]
    extracted_df = dataframe.iloc[:, target_indices]
    new_column_names = ['検討項目分類','課題番号','記入日','起票者','対象ドキュメント','完了予定日','優先度','難易度','Ｇ間調整','他Ｇヒアリング有無','完了予定の見通し','着手日','完了日','対応者','完了']
    if len(new_column_names) != len(extracted_df.columns): raise ValueError("抽出された列数と新しい列名の数が一致しません。")
    extracted_df.columns = new_column_names
    return extracted_df

def get_unique_values_from_column_with_start_row(dataframe: pd.DataFrame, column_index: int, start_row_index: int = 0) -> list:
    """
    DataFrameの指定列から、NaNや空文字列を除いた一意な値のリストを取得します。

    Args:
        dataframe (pd.DataFrame): 元データフレーム。
        column_index (int): 対象列インデックス。
        start_row_index (int, optional): 開始行インデックス。デフォルトは0。

    Returns:
        list: 一意な値のリスト。範囲外の場合は空リスト。
    """
    if not (0 <= column_index < len(dataframe.columns) and 0 <= start_row_index < len(dataframe)):
        print(f"警告: get_unique_valuesのインデックスが範囲外です。col:{column_index}, row:{start_row_index}")
        return []
    series = dataframe.iloc[start_row_index:, column_index]
    cleaned_series = series.dropna()
    if pd.api.types.is_string_dtype(cleaned_series):
        cleaned_series = cleaned_series[cleaned_series.astype(str).str.strip() != '']
    return cleaned_series.unique().tolist()

def load_data_to_dataframe(file_path: str, start_row: int, sheet_name: Union[str, int] = 0, usecols: Optional[str] = None) -> pd.DataFrame:
    """
    指定したExcelファイルまたはCSVファイルからデータを読み込み、DataFrameとして返します。

    Args:
        file_path (str): 読み込むファイルのパス。
        start_row (int): 読み飛ばす行数（0始まり）。
        sheet_name (Union[str, int], optional): シート名またはインデックス（Excelの場合）。デフォルトは0。
        usecols (Optional[str], optional): 読み込む列範囲（例: 'A:D'）。デフォルトはNone。

    Returns:
        pd.DataFrame: 読み込んだデータフレーム。

    Raises:
        FileNotFoundError: ファイルが存在しない場合。
        Exception: 読み込み時にエラーが発生した場合。
    """
    if not os.path.exists(file_path): raise FileNotFoundError(f"エラー: ファイル '{file_path}' が見つかりません。")
    try:
        return pd.read_excel(file_path, skiprows=start_row, sheet_name=sheet_name, usecols=usecols)
    except Exception as e: raise Exception(f"ファイル '{file_path}' の読み込み中にエラーが発生しました: {e}")

def analyze_document_tasks(target_df: pd.DataFrame) -> tuple:
    """
    文書ごとの未完了タスクを分析し、遅延状況などを算出します。

    Args:
        target_df (pd.DataFrame): 対象となるタスクデータフレーム。

    Returns:
        tuple: (遅延タスク数, 遅延タスクのリスト, 最大遅延日数, 最新の完了予定日文字列)
            遅延タスク数 (int): 期限を過ぎている未完了タスクの件数。
            遅延タスクのリスト (list): 期限を過ぎている未完了タスクの辞書リスト。
            最大遅延日数 (int): 最新の完了予定日と本日との差分（日数）。
            最新の完了予定日文字列 (str): 最大の完了予定日（'YYYY-MM-DD'形式または未設定時の文字列）。
    """
    if target_df.empty: return 0, [], 0, "タスク未設定"
    active_statuses = ['未着手', '着手中', '対応中']
    active_tasks_df = target_df[target_df['status'].isin(active_statuses)].copy()
    if active_tasks_df.empty: return 0, [], 0, "タスク完了済み"
    execution_date = pd.to_datetime('today').normalize()
    end_dates = pd.to_datetime(active_tasks_df['end_date_planned'], format='%Y-%m-%d', errors='coerce')
    max_date = end_dates.max()
    latest_date_str = max_date.strftime('%Y-%m-%d') if pd.notna(max_date) else "完了予定日未設定"
    delay_days = (max_date.date() - execution_date.date()).days if pd.notna(max_date) else 0
    is_overdue = end_dates < execution_date
    overdue_count = is_overdue.sum()
    overdue_tasks = active_tasks_df[is_overdue].to_dict('records')
    return overdue_count, overdue_tasks, delay_days, latest_date_str

def extract_delayed_tasks(input_df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    DataFrameから完了予定日を過ぎた未完了タスクと、完了予定日が未設定（NaT）の行を抽出します。

    Args:
        input_df (pd.DataFrame): 元データフレーム。

    Returns:
        tuple: (遅延タスクDataFrame, 完了予定日NaT行DataFrame)
            遅延タスクDataFrame (pd.DataFrame): 完了予定日を過ぎた未完了タスク。
            完了予定日NaT行DataFrame (pd.DataFrame): 完了予定日が未設定の行。
    """
    copy_df = input_df.copy()
    copy_df['完了予定日'] = pd.to_datetime(copy_df['完了予定日'], format='%Y-%m-%d', errors='coerce')
    today = pd.to_datetime('today').normalize()
    is_valid_date = copy_df['完了予定日'].notna()
    is_due = copy_df['完了予定日'] < today
    is_incomplete = copy_df['完了'] != '完了'
    delayed_tasks_df = copy_df[is_valid_date & is_due & is_incomplete].copy()
    nat_rows_df = copy_df[copy_df['完了予定日'].isna()]
    return delayed_tasks_df, nat_rows_df

def load_and_prepare_progress_data(online_file: str, other_file: str, config: dict, sheet_name: str) -> pd.DataFrame:
    """
    進捗データファイル（オンライン・その他）を読み込み、指定範囲の列を抽出・結合して返します。

    Args:
        online_file (str): オンライン設計書進捗ファイルのパス。
        other_file (str): その他設計書進捗ファイルのパス。
        config (dict): 読み込み設定辞書（開始行・列など）。
        sheet_name (str): 読み込むシート名。

    Returns:
        pd.DataFrame: 結合・整形済みの進捗データフレーム。
    """
    print("進捗データの読み込みと準備を開始...")
    online_df = load_data_to_dataframe(online_file, config['START_ROW_ONLINE'], sheet_name)
    other_df = load_data_to_dataframe(other_file, config['START_ROW'], sheet_name)
    extracted_online_df = extract_columns_by_range(online_df, config['START_COLUMN'], config['END_COLUMN'])
    extracted_other_df = extract_columns_by_range(other_df, config['START_COLUMN'], config['END_COLUMN'])
    all_df = pd.concat([extracted_online_df, extracted_other_df], axis=0, ignore_index=True)
    print("進捗データの準備が完了しました。")
    return all_df

def load_and_prepare_new_file_transfer_data(other_file: str, config: dict, sheet_name: str) -> pd.DataFrame:
    """
    新ファイル転送進捗データを読み込み、指定範囲の列を抽出して返します。

    Args:
        other_file (str): 新ファイル転送進捗ファイルのパス。
        config (dict): 読み込み設定辞書。
        sheet_name (str): 読み込むシート名。

    Returns:
        pd.DataFrame: 整形済みの新ファイル転送進捗データフレーム。
    """
    print("新ファイル転送進捗データの読み込みと準備を開始...")
    other_df = load_data_to_dataframe(other_file, config['START_ROW_ONLINE'], sheet_name)
    extracted_df = extract_columns_by_range(other_df, config['START_COLUMN_NEW_FILE_TRANSFER'], config['END_COLUMN_NEW_FILE_TRANSFER'])
    print("新ファイル転送進捗データの準備が完了しました。")
    return extracted_df

def load_and_prepare_issue_data(issue_file: str, config: dict) -> pd.DataFrame:
    """
    課題管理ファイルを読み込み、必要な列を抽出・整形して返します。

    Args:
        issue_file (str): 課題管理ファイルのパス。
        config (dict): 読み込み設定辞書。

    Returns:
        pd.DataFrame: 整形済みの課題データフレーム。
    """
    print("課題データの読み込みと準備を開始...")
    issue_df_raw = load_data_to_dataframe(issue_file, config['START_ROW_ISSUE_LIST'], '課題一覧')
    intermediate_df = extract_columns_by_range_issue(issue_df_raw)
    print("課題データの準備が完了しました。")
    return intermediate_df

def load_and_prepare_construction_data(file_path: str, config: dict, col_names: list) -> pd.DataFrame:
    """
    構築準備スケジュールデータを読み込み、前処理を行いDataFrameとして返します。

    Args:
        file_path (str): 構築準備スケジュールファイルのパス。
        config (dict): 読み込み設定辞書。
        col_names (list): 設定する列名リスト。

    Returns:
        pd.DataFrame: 前処理済みの構築準備データフレーム。

    Raises:
        ValueError: 読み込んだ列数と設定列名の数が一致しない場合。
    """
    print("構築準備スケジュールデータの読み込みと準備を開始...")
    df = load_data_to_dataframe(
        file_path, config['START_ROW_CONSTRUCTION_PREP'],
        config['CONSTRUCTION_PREP_SHEET_NAME'], config['USE_COLS_CONSTRUCTION_PREP']
    )
    df.dropna(how='all', inplace=True)
    if len(col_names) == len(df.columns): df.columns = col_names
    else: raise ValueError(f"読み込んだ列数({len(df.columns)})と設定列名の数({len(col_names)})が一致しません。")

    date_cols = ['開始予定日', '終了予定日', '開始実績日', '終了実績日']
    for col in date_cols: df[col] = pd.to_datetime(df[col], errors='coerce')

    numeric_cols = ['予定頁数', '実績頁数']
    for col in numeric_cols: df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    print("構築準備スケジュールデータの準備が完了しました。")
    return df

def extract_missing_page_tasks(online_file: str, other_file: str, config: dict, sheet_name: str) -> pd.DataFrame:
    """
    進捗ファイルから、完了済みかつ予定頁数または実績頁数が未設定のタスクを抽出します。

    Args:
        online_file (str): オンライン設計書進捗ファイルのパス。
        other_file (str): その他設計書進捗ファイルのパス。
        config (dict): 読み込み設定辞書。
        sheet_name (str): 読み込むシート名。

    Returns:
        pd.DataFrame: 頁数または実績未設定の完了タスクのみを含むDataFrame。
    """
    print("頁数/実績未設定の完了タスク抽出を開始...")
    combined_df = load_and_prepare_progress_data(online_file, other_file, config, sheet_name)
    target_processes = {
        "清書(PullRequest依頼⇒修正含む)", "設計書執筆（骨子への下書き）", "清書(PullRequest依頼まで)",
        "設計書執筆（別紙作成下書き）", "設計書執筆（MarkDown清書PullRequest依頼まで）&確認結果修正"
    }
    condition1 = combined_df['process_name'].isin(target_processes)
    condition2 = combined_df['status'] == '完了'
    page_planned_is_na = pd.to_numeric(combined_df['page_planned'], errors='coerce').isna()
    page_actual_is_na = pd.to_numeric(combined_df['page_actual'], errors='coerce').isna()
    condition3 = page_planned_is_na | page_actual_is_na
    filtered_df = combined_df[condition1 & condition2 & condition3].copy()
    print(f"頁数/実績未設定の完了タスクを {len(filtered_df)} 件抽出しました。")
    return filtered_df

def summarize_document_progress(progress_df: pd.DataFrame, process_definitions: dict, include_finalize: bool = True) -> pd.DataFrame:
    """
    進捗データから、設計書ごとの進捗状況を集計します。

    Args:
        progress_df (pd.DataFrame): 進捗データフレーム。
        process_definitions (dict): プロセス定義（執筆・清書・レビュー等の分類）。
        include_finalize (bool, optional): 清書プロセスも集計する場合はTrue。デフォルトはTrue。

    Returns:
        pd.DataFrame: 設計書ごとの進捗集計結果データフレーム。
    """
    print(f"進捗状況の集計を開始 (清書集計: {include_finalize})...")
    document_list = get_unique_values_from_column_with_start_row(progress_df, 0, 0)
    row_list = []
    for document_name in document_list:
        per_doc_df = progress_df[progress_df['document_name'] == document_name]
        per_doc_writing_df = per_doc_df[per_doc_df['process_name'].isin(process_definitions['writing'])]
        per_doc_team_review_df = per_doc_df[per_doc_df['process_name'].isin(process_definitions['team_review'])]
        per_doc_gl_review_df = per_doc_df[per_doc_df['process_name'].isin(process_definitions['gl_review'])]
        per_doc_psl_review_df = per_doc_df[per_doc_df['process_name'].isin(process_definitions['psl_review'])]
        total = len(per_doc_writing_df)
        status_counts = per_doc_writing_df['status'].value_counts()
        writing_ongoing = status_counts.get('着手中', 0) + status_counts.get('対応中', 0)
        writing_complete = status_counts.get('完了', 0)
        copy_writing_df = per_doc_writing_df.copy()
        copy_writing_df['page_planned_numeric'] = pd.to_numeric(copy_writing_df['page_planned'], errors='coerce')
        copy_writing_df['page_actual_numeric'] = pd.to_numeric(copy_writing_df['page_actual'], errors='coerce')
        total_pages_planned = copy_writing_df['page_planned_numeric'].sum()
        total_pages_actual = copy_writing_df['page_actual_numeric'].sum()
        writing_overdue, _, delay_date_count, latest_expected_completion_date = analyze_document_tasks(copy_writing_df)
        new_row = {'設計書名': document_name,'執筆全量': total, '執筆着手中': writing_ongoing, '執筆完了': writing_complete,'執筆消化率': f"{(writing_complete / total) if total != 0 else 0.0:.1%}",'執筆遅延': writing_overdue, '執筆完了予定日': latest_expected_completion_date, '執筆完了予定日に対する遅れ': delay_date_count,'執筆予定頁数': total_pages_planned, '執筆実績頁数': total_pages_actual,'執筆頁消化率': f"{(total_pages_actual / total_pages_planned) if total_pages_planned != 0 else 0.0:.1%}",}
        if include_finalize:
            per_doc_finalize_df = per_doc_df[per_doc_df['process_name'].isin(process_definitions['finalize'])]
            status_counts_finalize = per_doc_finalize_df['status'].value_counts()
            finalize_ongoing = status_counts_finalize.get('着手中', 0)
            finalize_complete = status_counts_finalize.get('完了', 0)
            copy_finalize_df = per_doc_finalize_df.copy()
            copy_finalize_df['page_planned_numeric'] = pd.to_numeric(copy_finalize_df['page_planned'], errors='coerce')
            copy_finalize_df['page_actual_numeric'] = pd.to_numeric(copy_finalize_df['page_actual'], errors='coerce')
            total_pages_planned_finalize = copy_finalize_df['page_planned_numeric'].sum()
            total_pages_actual_finalize = copy_finalize_df['page_actual_numeric'].sum()
            finalize_overdue, _, finalize_delay_date_count, finalize_latest_date = analyze_document_tasks(copy_finalize_df)
            new_row.update({'清書着手中': finalize_ongoing, '清書完了': finalize_complete,'清書消化率': f"{(finalize_complete / len(per_doc_finalize_df)) if len(per_doc_finalize_df) != 0 else 0.0:.1%}",'清書遅延': finalize_overdue, '清書完了予定日': finalize_latest_date, '清書完了予定日に対する遅れ': finalize_delay_date_count,'清書予定頁数': total_pages_planned_finalize, '清書実績頁数': total_pages_actual_finalize,'清書頁消化率': f"{(total_pages_actual_finalize / total_pages_planned_finalize) if total_pages_planned_finalize != 0 else 0.0:.1%}",})
        new_row.update({'T内RV中': per_doc_team_review_df['status'].value_counts().get('着手中', 0),'T内RV完了': per_doc_team_review_df['status'].value_counts().get('完了', 0),'T内RV全量': len(per_doc_team_review_df),'GL RV中': per_doc_gl_review_df['status'].value_counts().get('着手中', 0),'GL RV完了': per_doc_gl_review_df['status'].value_counts().get('完了', 0),'GL RV全量': len(per_doc_gl_review_df),'PSL RV中': per_doc_psl_review_df['status'].value_counts().get('着手中', 0),'PSL RV完了': per_doc_psl_review_df['status'].value_counts().get('完了', 0),'PSL RV全量': len(per_doc_psl_review_df),})
        row_list.append(new_row)
    print("進捗状況の集計が完了しました。")
    return pd.DataFrame(row_list)

def summarize_issues(issue_df: pd.DataFrame, custom_order: list) -> pd.DataFrame:
    """
    課題データから、設計書ごとの課題状況を集計・ソートします。

    Args:
        issue_df (pd.DataFrame): 課題データフレーム。
        custom_order (list): 設計書名の表示順リスト。

    Returns:
        pd.DataFrame: 設計書ごとの課題集計・ソート済みデータフレーム。
    """
    print("課題状況の集計を開始...")
    document_list_issue = get_unique_values_from_column_with_start_row(issue_df, 4, 0)
    issue_list = []
    for document_name in document_list_issue:
        per_doc_df = issue_df[issue_df['対象ドキュメント'] == document_name].copy()
        per_doc_df['完了'] = per_doc_df['完了'].str.replace(r'\s+', '', regex=True)
        per_doc_incomplete_df = per_doc_df[per_doc_df['完了'] != '完了']
        overdue_task_df, nat_rows_df = extract_delayed_tasks(per_doc_incomplete_df)
        new_row = {'設計書名': document_name,'設計書ごと課題全量': len(per_doc_df),'課題着手中': len(per_doc_incomplete_df[per_doc_incomplete_df['完了'].isin(['ヒアリングシート・QA発行Ｇ間資料待ち', '検討中'])]),'完了済み': len(per_doc_df) - len(per_doc_incomplete_df),'遅延': len(overdue_task_df['課題番号'].unique()),'完了予定日空白課題リスト': nat_rows_df['課題番号'].tolist(),'完了予定日超過課題リスト': overdue_task_df['課題番号'].unique().tolist(),'完了列空白課題リスト': per_doc_incomplete_df[per_doc_incomplete_df['完了'].isna()]['課題番号'].tolist(),}
        issue_list.append(new_row)
    result_issue_df = pd.DataFrame(issue_list)
    existing_values = result_issue_df['設計書名'].unique()
    filtered_order = [item for item in custom_order if item in existing_values]
    result_issue_df['設計書名'] = pd.Categorical(result_issue_df['設計書名'], categories=filtered_order, ordered=True)
    issue_df_sorted = result_issue_df.sort_values('設計書名')
    print("課題状況の集計が完了しました。")
    return issue_df_sorted

def process_construction_prep_data(prep_df: pd.DataFrame, kukan_to_phases_map: dict) -> pd.DataFrame:
    """
    構築準備データを区分ごとにグループ化し、集計します。

    Args:
        prep_df (pd.DataFrame): 構築準備データフレーム。
        kukan_to_phases_map (dict): 区分ごとのフェーズ定義マッピング。

    Returns:
        pd.DataFrame: 区分・対象・フェーズごとに集計された構築準備データフレーム。

    Raises:
        ValueError: 未定義の作業項目が存在する場合。
    """
    print("構築準備データの集計処理を開始...")
    all_grouped_dfs = []

    for kukan, phase_definitions in kukan_to_phases_map.items():
        kukan_df = prep_df[prep_df['区分'] == kukan].copy()
        if kukan_df.empty:
            print(f"INFO: 区分「{kukan}」のデータが見つかりませんでした。")
            continue

        item_to_phase_map = {item.strip(): phase for phase, items in phase_definitions.items() for item in items}
        kukan_df['フェーズ'] = kukan_df['作業項目'].str.strip().map(item_to_phase_map)

        unmapped_items = kukan_df[kukan_df['フェーズ'].isna()]['作業項目'].unique()
        if len(unmapped_items) > 0:
            unmapped_str = ", ".join(f"'{item}'" for item in unmapped_items)
            raise ValueError(f"区分「{kukan}」で未定義の作業項目が検出されました: {unmapped_str}")

        agg_rules = {'開始予定日': 'min', '終了予定日': 'max', '予定頁数': 'max', '実績頁数': 'min'}
        grouped = kukan_df.groupby(['区分', '対象', 'フェーズ']).agg(agg_rules)
        all_grouped_dfs.append(grouped)

    if not all_grouped_dfs:
        print("WARNING: 処理対象となる区分のデータがありませんでした。")
        return pd.DataFrame()

    final_df = pd.concat(all_grouped_dfs)
    print("構築準備データの集計処理が完了しました。")
    return final_df

def read_latest_past_totals(pattern: str) -> Dict[str, pd.DataFrame]:
    """
    昨日以前の最新レポートファイルを探し、存在する全シートから「合計」行を読み込みます。

    Args:
        pattern (str): ファイル検索パターン（例: '*-進捗課題状況.xlsx'）。

    Returns:
        Dict[str, pd.DataFrame]: シート名をキー、「合計」行のみ抽出したDataFrameを値とする辞書。
                                 ファイルが見つからない場合は空辞書を返します。

    Raises:
        例外発生時はエラーメッセージを出力し、空辞書を返します。
    """
    today = date.today()
    all_files = glob.glob(pattern)
    file_date_list = []
    for f in all_files:
        if m := re.search(r'(\d{4}-\d{2}-\d{2})', f):
            try:
                file_dt = datetime.strptime(m.group(1), "%Y-%m-%d").date()
                if file_dt < today: file_date_list.append((file_dt, f))
            except ValueError: continue
    if not file_date_list:
        print("⚠️ 読み込み対象となる昨日以前のファイルが見つかりませんでした。")
        return {}

    latest_file = max(file_date_list, key=lambda x: x[0])[1]
    print(f"✅ 過去レポート読み込み: {latest_file}")

    try:
        previous_dfs = pd.read_excel(latest_file, sheet_name=None)
        return {
            sheet_name: df[df.iloc[:, 0] == '合計'].copy()
            for sheet_name, df in previous_dfs.items()
            if not df.empty and df.iloc[:, 0].astype(str).str.contains('合計').any()
        }
    except Exception as e:
        print(f"❌ 過去レポートの読み込み中にエラーが発生しました: {e}")
        return {}

# ==============================================================================
# Excel出力関連クラス (リファクタリング適用)
# ==============================================================================

class SheetWriter:
    """Excelシート書き込み処理の設計図（基底クラス）"""
    def __init__(self, workbook: xlsxwriter.Workbook, config: ReportConfig):
        self.workbook = workbook
        self.config = config

    def write(self, sheet_name: str, data: pd.DataFrame, previous_total: Optional[pd.DataFrame] = None):
        raise NotImplementedError("This method should be implemented by subclasses.")

class StandardSheetWriter(SheetWriter):
    """標準的なフォーマット（データ+合計行+前回比）でシートを書き込むクラス"""
    def write(self, sheet_name: str, data: pd.DataFrame, previous_total: Optional[pd.DataFrame] = None):
        print(f"「{sheet_name}」を標準フォーマットで書き込み中...")
        worksheet = self.workbook.add_worksheet(sheet_name)
        sheet_config = self.config.SHEET_STYLES.get(sheet_name, {})
        if not sheet_config:
            print(f"WARNING: シート「{sheet_name}」のスタイル設定が見つかりません。")
            return

        formats = self._apply_formats(worksheet, sheet_config)

        # 日付列を変換
        date_cols_indices = sheet_config.get("date_cols", [])
        if date_cols_indices:
            date_cols_names = [data.columns[i] for i in date_cols_indices if i < len(data.columns)]
            preserved_strings = {"タスク未設定", "タスク完了済み", "完了予定日未設定"}
            def smart_date_converter(value):
                return value if value in preserved_strings or pd.isna(value) else pd.to_datetime(value, errors='coerce')
            for col in date_cols_names:
                if col in data.columns: data[col] = data[col].apply(smart_date_converter)

        # ヘッダー書き込み
        for col_idx, col_name in enumerate(data.columns):
            worksheet.write(0, col_idx, col_name, formats['header'])

        # データ書き込み
        highlight_config = sheet_config.get("highlight_rows")
        highlight_col = highlight_config["column"] if highlight_config else None
        highlight_vals = set(highlight_config["values"]) if highlight_config else set()
        for row_idx, row_data in enumerate(data.itertuples(index=False)):
            excel_row_idx = row_idx + 1
            is_highlight_row = highlight_col and getattr(row_data, highlight_col, None) in highlight_vals
            for col_idx, cell_value in enumerate(row_data):
                is_date_col = col_idx in date_cols_indices
                cell_format = formats['gray_bkg_date'] if is_highlight_row and is_date_col else \
                              formats['gray_bkg'] if is_highlight_row else \
                              formats['date'] if is_date_col else formats['default']
                if not pd.api.types.is_scalar(cell_value): worksheet.write_string(excel_row_idx, col_idx, str(cell_value), cell_format)
                elif pd.isna(cell_value) or (isinstance(cell_value, pd.Timestamp) and pd.isna(cell_value)): worksheet.write_blank(excel_row_idx, col_idx, None, cell_format)
                else: worksheet.write(excel_row_idx, col_idx, cell_value, cell_format)

        # 合計行と前回比行を追加
        num_data_rows = len(data)
        if "sum_cols" in sheet_config:
            total_row_idx = self._add_total_row(worksheet, num_data_rows, formats, sheet_config)
            self._add_comparison_rows(worksheet, total_row_idx, previous_total, formats, sheet_config)

    def _apply_formats(self, worksheet, sheet_config: dict) -> dict:
        for col_range, width in sheet_config.get("column_widths", {}).items():
            worksheet.set_column(col_range, width)
        base_props = {'valign': 'vcenter', 'align': 'center', 'text_wrap': True}
        worksheet.set_row(0, 41)
        return {
            'header': self.workbook.add_format({**base_props, 'bold': True, 'bg_color': '#DDEBF7'}),
            'default': self.workbook.add_format(base_props), 'bold': self.workbook.add_format({**base_props, 'bold': True}),
            'percent': self.workbook.add_format({**base_props, 'num_format': '0.0%'}),
            'percent_detailed': self.workbook.add_format({**base_props, 'num_format': '0.000%'}),
            'gray_bkg': self.workbook.add_format({**base_props, 'bg_color': '#D9D9D9'}),
            'total': self.workbook.add_format({**base_props, 'bold': True, 'bg_color': '#E2F0D9'}),
            'total_percent': self.workbook.add_format({**base_props, 'bold': True, 'bg_color': '#E2F0D9', 'num_format': '0.0%'}),
            'date': self.workbook.add_format({**base_props, 'num_format': 'yyyy-mm-dd'}),
            'gray_bkg_date': self.workbook.add_format({**base_props, 'bg_color': '#D9D9D9', 'num_format': 'yyyy-mm-dd'}),
        }

    def _add_total_row(self, worksheet, num_data_rows: int, formats: dict, sheet_config: dict) -> int:
        total_row_idx = num_data_rows + 1
        total_row_excel_num = total_row_idx + 1
        worksheet.set_row(total_row_idx, None, formats['total'])
        worksheet.write(total_row_idx, 0, '合計', formats['total'])
        for col_idx in sheet_config.get("sum_cols", []):
            col_letter = xlsxwriter.utility.xl_col_to_name(col_idx)
            formula = f'=SUM({col_letter}2:{col_letter}{num_data_rows + 1})'
            worksheet.write_formula(total_row_idx, col_idx, formula, formats['total'])
        for num_idx, den_idx, out_idx in sheet_config.get("ratio_cols", []):
            num_col, den_col = xlsxwriter.utility.xl_col_to_name(num_idx), xlsxwriter.utility.xl_col_to_name(den_idx)
            formula = f'=IFERROR({num_col}{total_row_excel_num}/{den_col}{total_row_excel_num},0)'
            worksheet.write_formula(total_row_idx, out_idx, formula, formats['total_percent'])
        return total_row_idx

    def _add_comparison_rows(self, worksheet, total_row_idx: int, previous_total_row: pd.DataFrame, formats: dict, sheet_config: dict):
        if previous_total_row is None or previous_total_row.empty:
            print(f"INFO: シート '{worksheet.name}' の前回集計データが見つからないため、前回比の行はスキップします。")
            return
        diff_row_idx, prev_total_row_idx = total_row_idx + 1, total_row_idx + 2
        total_row_excel_num, prev_total_row_excel_num = total_row_idx + 1, prev_total_row_idx + 1
        worksheet.write(diff_row_idx, 0, '前回比', formats['bold'])
        worksheet.write(prev_total_row_idx, 0, '前回集計の合計', formats['bold'])
        percent_cols = sheet_config.get("percent_cols", [])
        prev_values = previous_total_row.iloc[0]
        for i, value in enumerate(prev_values):
            if i > 0 and pd.api.types.is_number(value) and np.isfinite(value):
                 cell_format = formats['percent_detailed'] if i in percent_cols else formats['default']
                 worksheet.write(prev_total_row_idx, i, value, cell_format)
        for col_idx in sheet_config.get("diff_cols", []):
            col_letter = xlsxwriter.utility.xl_col_to_name(col_idx)
            formula = f'={col_letter}{total_row_excel_num}-{col_letter}{prev_total_row_excel_num}'
            cell_format = formats['percent_detailed'] if col_idx in percent_cols else formats['default']
            worksheet.write_formula(diff_row_idx, col_idx, formula, cell_format)

class ConstructionPrepSheetWriter(SheetWriter):
    """「構築準備」シートをカスタムレイアウトで書き込むクラス"""
    def write(self, sheet_name: str, data: pd.DataFrame, previous_total: Optional[pd.DataFrame] = None):
        print(f"「{sheet_name}」シートのカスタム出力を開始...")
        worksheet = self.workbook.add_worksheet(sheet_name)
        title_format = self.workbook.add_format({'bold': True, 'font_size': 18, 'valign': 'vcenter'})
        phase_header_format = self.workbook.add_format({'bold': True, 'bg_color': '#DDEBF7', 'align': 'center', 'valign': 'vcenter', 'border': 1})
        sub_header_format = self.workbook.add_format({'bold': True, 'bg_color': '#E2F0D9', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})
        target_format = self.workbook.add_format({'bold': True, 'align': 'left', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})
        default_format = self.workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
        date_format = self.workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': 'yyyy-mm-dd'})

        sub_headers = ['開始予定日', '終了予定日', '残り日数', '予定頁数', '実績頁数']
        cols_per_phase = len(sub_headers)

        if data.empty or '区分' not in data.index.names:
            worksheet.write(0, 0, "処理対象の区分のデータが見つかりませんでした。")
            print(f"「{sheet_name}」シートのデータが空のため、出力をスキップしました。")
            return

        worksheet.set_column(0, 0, 35)

        current_row = 0; today = datetime.now()
        all_kukan = data.index.get_level_values('区分').unique()

        for kukan in all_kukan:
            kukan_phases = list(self.config.KUKAN_TO_PHASES_MAP.get(kukan, {}).keys())
            if not kukan_phases: continue

            total_cols = len(kukan_phases) * cols_per_phase
            for i in range(total_cols): worksheet.set_column(i + 1, i + 1, 12)

            worksheet.merge_range(current_row, 0, current_row, total_cols, kukan, title_format)
            worksheet.set_row(current_row, 24); current_row += 2

            header_row, sub_header_row, data_start_row = current_row, current_row + 1, current_row + 2

            for i, phase in enumerate(kukan_phases):
                start_col, end_col = 1 + i * cols_per_phase, 1 + (i + 1) * cols_per_phase - 1
                if start_col <= end_col: worksheet.merge_range(header_row, start_col, header_row, end_col, phase, phase_header_format)
                for j, sub_header in enumerate(sub_headers):
                    worksheet.write(sub_header_row, start_col + j, sub_header, sub_header_format)

            kukan_df = data[data.index.get_level_values('区分') == kukan]
            all_targets = kukan_df.index.get_level_values('対象').unique()

            for i, target in enumerate(all_targets):
                row = data_start_row + i
                worksheet.write(row, 0, target, target_format)
                for j, phase in enumerate(kukan_phases):
                    start_col = 1 + j * cols_per_phase
                    try:
                        record = kukan_df.loc[(kukan, target, phase)]
                        start_date, end_date = record['開始予定日'], record['終了予定日']

                        if pd.notna(start_date): worksheet.write(row, start_col, start_date, date_format)
                        else: worksheet.write_blank(row, start_col, None, default_format)

                        if pd.notna(end_date):
                            worksheet.write(row, start_col + 1, end_date, date_format)
                            days_left = (end_date - today).days
                            worksheet.write(row, start_col + 2, days_left, default_format)
                        else:
                            worksheet.write_blank(row, start_col + 1, None, default_format)
                            worksheet.write_blank(row, start_col + 2, None, default_format)

                        worksheet.write(row, start_col + 3, record['予定頁数'], default_format)
                        worksheet.write(row, start_col + 4, record['実績頁数'], default_format)
                    except KeyError:
                        for k in range(cols_per_phase): worksheet.write_blank(row, start_col + k, None, default_format)
            current_row = data_start_row + len(all_targets) + 2
        print(f"「{sheet_name}」シートのカスタム出力が完了しました。")

def save_summaries_to_excel(summaries: Dict[str, pd.DataFrame], previous_totals_map: Dict[str, pd.DataFrame], execution_date: date, config: ReportConfig):
    """複数の集計結果DataFrameを1つのExcelファイルにシート分けして保存します。"""
    output_filename = f"{execution_date}-進捗課題状況.xlsx"
    print(f"集計結果を'{output_filename}'に出力中...")

    with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
        workbook = writer.book

        # シート名に応じて使用する書き込みクラスを定義
        writer_map = {
            config.SHEET_NAMES["progress"]: StandardSheetWriter(workbook, config),
            config.SHEET_NAMES["issue"]: StandardSheetWriter(workbook, config),
            config.SHEET_NAMES["new_transfer"]: StandardSheetWriter(workbook, config),
            config.SHEET_NAMES["missing_page"]: StandardSheetWriter(workbook, config),
            config.SHEET_NAMES["construction_prep"]: ConstructionPrepSheetWriter(workbook, config)
        }

        for sheet_name, data in summaries.items():
            if data is None or data.empty:
                print(f"INFO: シート '{sheet_name}' のデータが空のため、書き込みをスキップします。")
                continue

            if sheet_name in writer_map:
                writer_instance = writer_map[sheet_name]
                previous_total = previous_totals_map.get(sheet_name)
                writer_instance.write(sheet_name, data.copy(), previous_total)
            else:
                print(f"WARNING: シート '{sheet_name}' に対応する書き込みクラスが見つかりません。")

    print("Excelファイルの出力が完了しました。")

# ==============================================================================
# メイン処理
# ==============================================================================
def main():
    """スクリプトのエントリーポイント。"""
    config = ReportConfig()
    execution_date = datetime.now().date()

    previous_totals = read_latest_past_totals(pattern='*-進捗課題状況.xlsx')
    all_summaries = {}

    try:
        progress_df = load_and_prepare_progress_data(config.FILES['online_progress'], config.FILES['other_progress'], config.LOAD_CONFIG, config.LOAD_CONFIG['PROGRESS_SHEET_NAME'])
        all_summaries[config.SHEET_NAMES["progress"]] = summarize_document_progress(progress_df, config.PROCESS_DEFINITIONS, include_finalize=True)

        new_transfer_df = load_and_prepare_new_file_transfer_data(config.FILES['other_progress'], config.LOAD_CONFIG, config.LOAD_CONFIG['NEW_FILE_TRANSFER_PROGRESS_SHEET_NAME'])
        all_summaries[config.SHEET_NAMES["new_transfer"]] = summarize_document_progress(new_transfer_df, config.PROCESS_DEFINITIONS, include_finalize=False)

        issue_df = load_and_prepare_issue_data(config.FILES['issue_list'], config.LOAD_CONFIG)
        all_summaries[config.SHEET_NAMES["issue"]] = summarize_issues(issue_df, config.DOCUMENT_ORDER)

        all_summaries[config.SHEET_NAMES["missing_page"]] = extract_missing_page_tasks(config.FILES['online_progress'], config.FILES['other_progress'], config.LOAD_CONFIG, config.LOAD_CONFIG['PROGRESS_SHEET_NAME'])

        raw_prep_df = load_and_prepare_construction_data(
            config.FILES['construction_schedule'], config.LOAD_CONFIG, config.COL_NAMES_CONSTRUCTION_PREP
        )
        all_summaries[config.SHEET_NAMES["construction_prep"]] = process_construction_prep_data(raw_prep_df, config.KUKAN_TO_PHASES_MAP)

    except (FileNotFoundError, ValueError, IndexError) as e:
        print(f"\nエラー: データ処理中に問題が発生しました。入力ファイルや設定を確認してください。")
        print(f"詳細: {e}")
        return

    save_summaries_to_excel(
        summaries=all_summaries,
        previous_totals_map=previous_totals,
        execution_date=execution_date,
        config=config
    )

if __name__ == "__main__":
    main()