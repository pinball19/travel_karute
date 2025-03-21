/**
 * 一体型カルテのテンプレート
 * すべてのセクションを1つのテンプレートにまとめる
 */
const UNIFIED_TEMPLATE = [
  // 基本情報セクション
  ['◆ 基本情報', '', '', '', '', '', '', ''],
  ['【団体ナビ成約カルテ】', '', '担当；', '', '記入日;', '', '個人通N0;', ''],
  ['カルテNo', '', '名前', '', '団体名', '', '電話', ''],
  ['宿泊日', '', '泊数', '', '合計\n人数', '', '成約日', ''],
  ['出発地', '', '⇒', '行先', '', '', '', ''],
  ['', '', '', '', '', '', '', ''],
  
  // 入金情報セクション
  ['◆ 入金情報', '', '', '', '', '', '', ''],
  ['入金予定日', '入金場所', '入金予定額', '入金日', '入金額', '備考', 'チェック', ''],
  ['', '', '', '', '', '', '', ''],
  ['', '', '', '', '', '', '', ''],
  ['', '', '', '', '', '', '', ''],
  ['', '', '', '', '', '', '', ''],
  ['A；入金合計', '', '', '', '', '', '', ''],
  ['', '', '', '', '', '', '', ''],
  
  // 支払情報セクション
  ['◆ 支払情報', '', '', '', '', '', '', ''],
  ['利用日', '手配先名；該当に〇を付ける', '電話/FAX', '担当者', '支払予定日', '支払金額', 'チェック', ''],
  ['', '', '', '', '', '', '', ''],
  ['', '', '', '', '', '', '', ''],
  ['', '', '', '', '', '', '', ''],
  ['', '', '', '', '', '', '', ''],
  ['B；支払合計', '', '', '', '', '', '', ''],
  ['', '', '', '', '', '', '', ''],
  
  // 収支情報セクション
  ['◆ 収支情報', '', '', '', '', '', '', ''],
  ['報告日', '利益率', '利益額', '一人粗利', '旅行総額/A', '支払総額/B', '人数', '特補人数'],
  ['', '', '', '', '', '', '', ''],
  ['', '', '', '', '', '', '', ''],
  ['上席チェック', '何月/何日に利益申請', '', '', '', '', '', ''],
  ['', '', '', '', '', '', '', ''],
  ['', '', '', '', '', '', '', ''],
  ['', '', '', '', '', '', '', ''],
  ['', '', '', '', '', '', '', ''],
];

/**
 * 一体型のセル結合設定
 */
const UNIFIED_MERGED_CELLS = [
  // 基本情報セクション
  {row: 0, col: 0, rowspan: 1, colspan: 8}, // セクション見出し
  {row: 1, col: 0, rowspan: 1, colspan: 2}, // 【団体ナビ成約カルテ】
  
  // 入金情報セクション
  {row: 6, col: 0, rowspan: 1, colspan: 8}, // セクション見出し
  {row: 12, col: 0, rowspan: 1, colspan: 7}, // A；入金合計
  
  // 支払情報セクション
  {row: 14, col: 0, rowspan: 1, colspan: 8}, // セクション見出し
  {row: 20, col: 0, rowspan: 1, colspan: 7}, // B；支払合計
  
  // 収支情報セクション
  {row: 22, col: 0, rowspan: 1, colspan: 8}, // セクション見出し
  {row: 25, col: 0, rowspan: 1, colspan: 2}, // 上席チェック
];

/**
 * セクションごとのスタイル設定（行インデックス）
 */
const SECTION_STYLES = {
  // セクション見出し行
  sectionHeaders: [0, 6, 14, 22],
  
  // ヘッダー行（列名などが入る行）
  headerRows: [1, 2, 7, 15, 23, 25],
  
  // 合計行
  totalRows: [12, 20],
  
  // 空白の区切り行
  spacerRows: [5, 13, 21, 26, 27, 28, 29]
};

/**
 * カラム幅の設定
 */
const COLUMN_WIDTHS = [
  140, // 1列目：左側の項目名（幅広め・折り返し対応）
  120, // 2列目
  120, // 3列目
  120, // 4列目
  120, // 5列目
  120, // 6列目
  120, // 7列目
  120  // 8列目
];
