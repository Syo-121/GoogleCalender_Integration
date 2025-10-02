//グローバル設定
const SYNC_DAYS_PAST = 30;
const SYNC_DAYS_FUTURE = 180;
const SYNC_PROPERTY_KEY = 'gas.calendar.sync.sourceEventId.final.v2'; // 識別キーを更新

// UI設定
function onOpen() { /* ... */ }
function openSidebar() { /* ... */ }
function getCalendarList() { /* ... */ }
function getSettings() { /* ... */ }
function saveSettings(settings) { /* ... */ }

//スプシからプレフィックス設定を読み込む
function getPrefixMap() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('プレフィックス設定');
    if (!sheet) {
      return new Map(); // シートがなければ空のマップを返す
    }
    // A列とB列の2列だけを読み込む
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    const prefixMap = new Map();
    data.forEach(row => {
      const originalPrefix = row[0];
      const replacementText = row[1];
      if (originalPrefix && replacementText) {
        prefixMap.set(originalPrefix, replacementText);
      }
    });
    return prefixMap;
  } catch (e) {
    console.error('プレフィックス設定シートの読み込みに失敗しました。', e.message);
    return new Map(); // エラー時も空のマップを返す
  }
}

//タイトルの変換
function formatTitle(originalTitle, prefixMap) {
  if (!originalTitle) {
    return '予定あり';
  }

  const colonIndex = originalTitle.indexOf(':');

  // コロンがない場合は、一律「予定あり」
  if (colonIndex === -1) {
    return '予定あり';
  }

  // --- コロンがある場合の処理 ---

  // まず、プレフィックスルールに一致するかチェック
  for (const [prefix, replacement] of prefixMap.entries()) {
    if (originalTitle.startsWith(prefix)) {
      // ルールに一致したら、置換後のテキスト + 「予定あり」を返す
      return `${replacement} 予定あり`;
    }
  }

  // どのプレフィックスルールにも一致しなかった場合
  // 元のコロン前の部分をそのまま使い、「予定あり」を付ける
  const originalPrefix = originalTitle.substring(0, colonIndex).trim();
  return `${originalPrefix}: 予定あり`;
}

//カレンダーの同期
function syncAllCalendars() {
  const prefixMap = getPrefixMap();
  const settings = getSettings();
  if (!settings.sourceIds || !settings.targetId || settings.sourceIds.length === 0) {
    console.log('同期設定が未保存のためスキップしました。');
    return;
  }
  const { sourceIds: SOURCE_CALENDAR_IDS, targetId: TARGET_CALENDAR_ID } = settings;
  console.log(`【同期開始】集約先カレンダーID: ${TARGET_CALENDAR_ID}`);

  const startTime = new Date();
  startTime.setDate(startTime.getDate() - SYNC_DAYS_PAST);
  const endTime = new Date();
  endTime.setDate(endTime.getDate() + SYNC_DAYS_FUTURE);

  const syncedEventsMap = new Map();
  try {
    let pageToken;
    do {
      const response = Calendar.Events.list(TARGET_CALENDAR_ID, {
        timeMin: startTime.toISOString(),
        timeMax: endTime.toISOString(),
        pageToken: pageToken,
        maxResults: 2500
      });
      if (response.items) {
        response.items.forEach(event => {
          if (event.extendedProperties && event.extendedProperties.private && event.extendedProperties.private[SYNC_PROPERTY_KEY]) {
            const sourceEventId = event.extendedProperties.private[SYNC_PROPERTY_KEY];
            syncedEventsMap.set(sourceEventId, event);
          }
        });
      }
      pageToken = response.nextPageToken;
    } while (pageToken);
  } catch (e) {
    console.error(`集約先カレンダーのイベント取得に失敗しました: ${e.message}`);
    return;
  }
  console.log(`${syncedEventsMap.size}件の同期済みイベント情報を読み込みました。`);

  SOURCE_CALENDAR_IDS.forEach(sourceCalId => {
    console.log(`処理中: ${sourceCalId}`);
    try {
      let pageToken;
      do {
        const response = Calendar.Events.list(sourceCalId, {
          timeMin: startTime.toISOString(),
          timeMax: endTime.toISOString(),
          singleEvents: true,
          orderBy: 'startTime',
          pageToken: pageToken,
          maxResults: 2500,
        });
        if (response.items) {
          response.items.forEach(sourceEvent => {
            if (sourceEvent.status === 'cancelled') return;
            const sourceEventId = sourceEvent.id;
            const targetEvent = syncedEventsMap.get(sourceEventId);

            if (targetEvent) {
              const expectedTitle = formatTitle(sourceEvent.summary, prefixMap);
              const targetStart = new Date(targetEvent.start.dateTime || targetEvent.start.date).getTime();
              const sourceStart = new Date(sourceEvent.start.dateTime || sourceEvent.start.date).getTime();
              const targetEnd = new Date(targetEvent.end.dateTime || targetEvent.end.date).getTime();
              const sourceEnd = new Date(sourceEvent.end.dateTime || sourceEvent.end.date).getTime();

              // タイトル、開始時刻、終了時刻、場所、説明、色のいずれかが違うかをチェック
              const needsUpdate = (
                targetEvent.summary !== expectedTitle ||
                targetStart !== sourceStart ||
                targetEnd !== sourceEnd ||
                (targetEvent.location || '') !== (sourceEvent.location || '') ||
                (targetEvent.description || '') !== (sourceEvent.description || '') ||
                (targetEvent.colorId || null) !== (sourceEvent.colorId || null)
              );

              if (needsUpdate) {
                updateEvent(targetEvent, sourceEvent, TARGET_CALENDAR_ID, prefixMap);
                console.log(`  更新: ${sourceEvent.summary} (理由: 内容の不一致)`);
              }
              syncedEventsMap.delete(sourceEventId);
            } else { // --- 新規作成処理 ---
              createEvent(sourceEvent, TARGET_CALENDAR_ID, prefixMap);
              console.log(`  作成: ${sourceEvent.summary}`);
            }
          });
        }
        pageToken = response.nextPageToken;
      } while (pageToken);
    } catch (e) {
      console.error(`同期元カレンダー[${sourceCalId}]の処理中にエラー: ${e.message}`);
    }
  });

  syncedEventsMap.forEach(eventToDelete => {
    try {
      Calendar.Events.remove(TARGET_CALENDAR_ID, eventToDelete.id);
      console.log(`  削除: ${eventToDelete.summary}`);
    } catch (e) {
      console.error(`イベント[${eventToDelete.summary}]の削除中にエラー: ${e.message}`);
    }
  });
  console.log("【同期完了】");
}


//イベントデータ作成
function buildEventResource(sourceEvent, prefixMap, existingProperties) {
  // 必要なプロパティだけを丁寧に取り出して、新しいオブジェクトを作成する
  const resource = {
    summary: formatTitle(sourceEvent.summary, prefixMap),
    description: sourceEvent.description || '',
    location: sourceEvent.location || '',
    start: {
      dateTime: sourceEvent.start.dateTime,
      date: sourceEvent.start.date,
      timeZone: sourceEvent.start.timeZone
    },
    end: {
      dateTime: sourceEvent.end.dateTime,
      date: sourceEvent.end.date,
      timeZone: sourceEvent.end.timeZone
    },
    colorId: sourceEvent.colorId,
    // 更新時は既存のプロパティを維持し、作成時は新しいプロパティを付与する
    extendedProperties: existingProperties || {
      private: {
        [SYNC_PROPERTY_KEY]: sourceEvent.id
      }
    }
  };
  return resource;
}


// 集約先のカレンダーにイベントを作成する
function createEvent(sourceEvent, targetCalendarId, prefixMap) {
  // 安全なイベントデータを作成
  const newEventResource = buildEventResource(sourceEvent, prefixMap, null);
  try {
    Calendar.Events.insert(newEventResource, targetCalendarId);
  } catch (e) {
    console.error(`イベント作成エラー for "${newEventResource.summary}": ${e.message}`);
  }
}

// 既存のイベントの更新
function updateEvent(targetEvent, sourceEvent, targetCalendarId, prefixMap) {
  // 安全なイベントデータを作成（既存の識別情報は引き継ぐ）
  const updatedEventResource = buildEventResource(sourceEvent, prefixMap, targetEvent.extendedProperties);
  try {
    Calendar.Events.update(updatedEventResource, targetCalendarId, targetEvent.id);
  } catch (e) {
    console.error(`イベント更新エラー for "${updatedEventResource.summary}": ${e.message}`);
  }
}

// 設定を開く
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('カレンダー同期設定')
    .addItem('設定を開く', 'openSidebar')
    .addToUi();
}

// サイドバーを開く
function openSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('カレンダー同期設定');
  SpreadsheetApp.getUi().showSidebar(html);
}

//　カレンダーの一覧を取得
function getCalendarList() {
  return CalendarApp.getAllCalendars().map(cal => ({
    id: cal.getId(),
    name: cal.getName(),
    isOwned: cal.isMyPrimaryCalendar() || cal.isOwnedByMe(),
  }));
}

// 設定を取得
function getSettings() {
  const properties = PropertiesService.getUserProperties();
  const settings = properties.getProperty('calendarSyncSettings');
  return settings ? JSON.parse(settings) : {};
}

function saveSettings(settings) {
  PropertiesService.getUserProperties().setProperty('calendarSyncSettings', JSON.stringify(settings));
}
