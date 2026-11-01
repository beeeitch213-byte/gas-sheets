// ==== 共通: 休場日(祝日)の判定 ====
// 株式情報(stock-sheet/main)と損益管理(pl-sheet)の両方で使う。リポジトリの common/祝日.js が元で、
// 自動デプロイ(.github/workflows/clasp-push.yml)が clasp push の前に、両方のプロジェクトへ「共通_祝日.js」としてコピーする。
// 各プロジェクトのフォルダにあるコピーは直さない(次のデプロイで上書きされる)。直すときはこのファイルを直す。
//
// 休場日: 土日、国民の祝日・休日(振替休日・国民の休日を含む)、年末年始(1/1〜1/3・12/31、東証の休業日)。
// 祝日は内閣府の「国民の祝日」CSV を正とする。取得した内容はスクリプトプロパティに年ごとに保存し、
// JP_HOLIDAY_REFRESH_DAYS_ 日ごとに取り直す。CSV が取れない・その年が載っていないときだけ、
// Google の「日本の祝日」カレンダーで判定する(七夕などの行事を除くため、祝日の名前で絞る)。
//
// 両プロジェクトの全ファイルとグローバルスコープを共有するので、名前は JP_HOLIDAY_ / JpHoliday / isJp で始め、
// 読み込み時に他のファイルの定数を使わない(ファイルの並び順に影響されないようにするため)。

const JP_HOLIDAY_CSV_URL_ = 'https://www8.cao.go.jp/chosei/shukujitsu/syukujitsu.csv';

const JP_HOLIDAY_REFRESH_DAYS_ = 30; // 保存した祝日を取り直す間隔(日)。法改正などでの変更を拾うため

const JP_HOLIDAY_PROP_PREFIX_ = 'JP_HOLIDAYS_'; // スクリプトプロパティのキー: JP_HOLIDAYS_2026 など

const JP_HOLIDAY_CALENDAR_ID_ = 'ja.japanese#holiday@group.v.calendar.google.com';

// 予備のカレンダー判定で、祝日として扱う名前(内閣府の CSV では振替休日・国民の休日の名前が「休日」)
const JP_HOLIDAY_FALLBACK_NAMES_ = [
  '元日', '成人の日', '建国記念の日', '天皇誕生日', '春分の日', '昭和の日',
  '憲法記念日', 'みどりの日', 'こどもの日', '海の日', '山の日', '敬老の日',
  '秋分の日', 'スポーツの日', '体育の日', '文化の日', '勤労感謝の日',
  '振替休日', '国民の休日', '休日'
];

// 1回の実行の中では、年ごとの祝日の集合と CSV の取得結果を使い回す
const jpHolidayYearCache_ = {};
let jpHolidayCsvCache_ = null; // { holidaysByYear: {2026: {'2026-01-01': '元日', ...}} } または { failed: true }

/** その日が土日・祝日・年末年始なら true(営業日かどうかは !isJpMarketHoliday_(date)) */
function isJpMarketHoliday_(date) {
  const dow = date.getDay();
  return dow === 0 || dow === 6 || isJpHolidayOrYearEnd_(date);
}

/** その日が祝日・休日か年末年始(1/1〜1/3・12/31)なら true。土日は見ない */
function isJpHolidayOrYearEnd_(date) {
  const month = date.getMonth() + 1;
  const day = date.getDate();
  if ((month === 1 && day <= 3) || (month === 12 && day === 31)) return true;
  return !!getJpHolidaysForYear_(date.getFullYear())[jpHolidayDateKey_(date)];
}

/** その年の祝日を {'yyyy-MM-dd': 名前} で返す */
function getJpHolidaysForYear_(year) {
  if (jpHolidayYearCache_[year]) return jpHolidayYearCache_[year];

  const props = PropertiesService.getScriptProperties();
  const key = JP_HOLIDAY_PROP_PREFIX_ + year;
  let saved = null;
  try { saved = JSON.parse(props.getProperty(key) || 'null'); } catch (e) {}
  const fresh = saved && Date.now() - saved.fetchedAt < JP_HOLIDAY_REFRESH_DAYS_ * 24 * 3600 * 1000;

  let holidays = fresh ? saved.holidays : null;
  if (!holidays) {
    const csv = fetchJpHolidayCsv_();
    if (csv.holidaysByYear && csv.holidaysByYear[year]) {
      holidays = csv.holidaysByYear[year];
      props.setProperty(key, JSON.stringify({ fetchedAt: Date.now(), holidays: holidays }));
    } else if (saved) {
      holidays = saved.holidays; // 取り直せなかったときは、古くても保存済みのものを使う
      console.warn(year + '年の祝日: 内閣府の CSV を取り直せなかったため、保存済みのものを使います');
    } else {
      holidays = getJpHolidaysFromCalendar_(year);
      console.warn(year + '年の祝日: 内閣府の CSV から取れなかったため、Google カレンダーで判定します' +
        (csv.failed ? '(CSV の取得に失敗)' : '(CSV にこの年が載っていない)'));
    }
  }
  jpHolidayYearCache_[year] = holidays;
  return holidays;
}

/** 内閣府の CSV を取得して年ごとに分ける(1回の実行で1回だけ取りに行く) */
function fetchJpHolidayCsv_() {
  if (jpHolidayCsvCache_) return jpHolidayCsvCache_;
  try {
    const res = UrlFetchApp.fetch(JP_HOLIDAY_CSV_URL_, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) throw new Error('HTTP ' + res.getResponseCode());
    const holidaysByYear = {};
    res.getContentText('Shift_JIS').split(/\r?\n/).forEach(function (line) {
      const m = line.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2}),(.*)$/); // 1行目の見出しなどは読み飛ばす
      if (!m) return;
      const dateKey = m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
      (holidaysByYear[m[1]] = holidaysByYear[m[1]] || {})[dateKey] = m[4].trim();
    });
    if (Object.keys(holidaysByYear).length === 0) throw new Error('祝日が1件も読み取れませんでした');
    jpHolidayCsvCache_ = { holidaysByYear: holidaysByYear };
  } catch (e) {
    console.warn('内閣府の祝日 CSV を取得できませんでした: ' + e);
    jpHolidayCsvCache_ = { failed: true };
  }
  return jpHolidayCsvCache_;
}

/** 予備: Google の「日本の祝日」カレンダーから、その年の祝日を {'yyyy-MM-dd': 名前} で返す(取れなければ空) */
function getJpHolidaysFromCalendar_(year) {
  const holidays = {};
  try {
    const cal = CalendarApp.getCalendarById(JP_HOLIDAY_CALENDAR_ID_);
    cal.getEvents(new Date(year, 0, 1), new Date(year + 1, 0, 1)).forEach(function (e) {
      const name = e.getTitle();
      if (JP_HOLIDAY_FALLBACK_NAMES_.indexOf(name) === -1) return;
      holidays[jpHolidayDateKey_(e.isAllDayEvent() ? e.getAllDayStartDate() : e.getStartTime())] = name;
    });
  } catch (e) {
    console.warn('祝日カレンダーも取得できませんでした。土日・年末年始だけを休場にします: ' + e);
  }
  return holidays;
}

function jpHolidayDateKey_(date) {
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd');
}
