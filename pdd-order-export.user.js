// ==UserScript==
// @name         拼多多買家訂單匯出 (增強版)
// @namespace    https://github.com/DSH/pdd-order-export
// @version      1.9.0
// @description  自動攔截拼多多網頁版買家訂單資料，自動載入訂單、日期範圍篩選、自選匯出欄位（記住選項），一鍵匯出 Excel(.xlsx)/CSV。
// @author       leolai
// @match        https://mobile.pinduoduo.com/*
// @match        https://mobile.yangkeduo.com/*
// @grant        unsafeWindow
// @grant        GM_setClipboard
// @grant        GM_registerMenuCommand
// @grant        GM_getValue
// @grant        GM_setValue
// @run-at       document-start
// @noframes
// ==/UserScript==

(function () {
  'use strict';

  /* ------------------------------------------------------------------ *
   *  常數 / 全域狀態
   * ------------------------------------------------------------------ */
  const NS = 'pdd-order-export';
  const ORDER = {};               // key = orderSn，value = normalized record
  const RAW = {};                 // key = orderSn，value = 原始物件（重整／換單位用）
  const SEEN = new Set();         // 已處理過嘅原始物件，避免重複
  let detectionMethod = '未偵測';
  let autoLoading = false;
  let findBudget = 1e6;           // 每次掃描嘅總節點預算（避免卡死）
  let interceptCount = 0;         // 網路攔截到的回應數（偵錯）
  let lastUrl = '';               // 最後攔截 URL（偵錯）

  const dbg = (...a) => { try { if (unsafeWindow.__pddDebug) console.log('[pdd-export]', ...a); } catch (e) {} };

  /* ------------------------------------------------------------------ *
   *  記住用戶選項（Tampermonkey 持久儲存，後備 localStorage）
   * ------------------------------------------------------------------ */
  const store = {
    get(k, def) {
      try { if (typeof GM_getValue === 'function') { const v = GM_getValue('pdd_' + k, undefined); if (v !== undefined) return v; } } catch (e) {}
      try { const v = localStorage.getItem('pdd_' + k); return v == null ? def : JSON.parse(v); } catch (e) {}
      return def;
    },
    set(k, v) {
      try { if (typeof GM_setValue === 'function') GM_setValue('pdd_' + k, v); } catch (e) {}
      try { localStorage.setItem('pdd_' + k, JSON.stringify(v)); } catch (e) {}
    },
  };

  /* ------------------------------------------------------------------ *
   *  多語言（繁體中文 / 簡體中文 / 英文）
   * ------------------------------------------------------------------ */
  const I18N = {
    'zh-HK': {
      'app.title': '拼多多訂單導出工具',
      'help.btn': '使用說明', 'close': '關閉', 'status.prefix': '狀態：',
      'theme.btn': '暗黑模式', 'theme.toDark': '切換至暗黑模式', 'theme.toLight': '切換至淺色模式',
      'help.title': '點樣用？', 'help.ok': '知道了', 'help.close': '關閉/返回',
      'help.body': `<ol>
        <li>打開 <b>拼多多網頁版</b> 並登入。</li>
        <li>入「我的訂單」→「<b>查看全部</b>」訂單列表頁面。</li>
        <li>揀<b>日期範圍</b>（從／到）。</li>
        <li>（可選）撳「選擇匯出欄位」勾你想匯出嘅項目。</li>
        <li>撳「<b>搜尋 &amp; 匯出 Excel</b>」：會自動載入訂單、按日期過濾、再匯出 Excel（含統計總覽）。</li>
        <li>載入期間可撳橙色「<b>停止搜尋</b>」隨時停。</li>
      </ol><p>金額會自動由「分」換算做「元」；你揀過嘅匯出欄位會記住。</p>`,
      'field.range': '日期範圍', 'field.from': '從', 'field.to': '到',
      'field.cols': '選擇匯出欄位（已選 {n} 項）▾', 'field.colsOpen': '收起', 'field.colsAll': '全選 / 全不選',
      'search': '搜尋 & 匯出 Excel', 'searchStop': '⏹ 停止搜尋',
      'note': '撳「搜尋」後會自動載入訂單，只匯出所揀日期範圍內嘅訂單。',
      'status.idle': '揀日期後撳「搜尋」',
      'status.searching': '搜尋中… 載入緊訂單', 'status.loading': '載入訂單中…',
      'status.none': '呢段日期無搵到訂單',
      'status.found': '搵到 {n} 筆，匯出中…', 'status.done': '完成，已匯出 {n} 筆',
      'status.stopped': '已停止（已搵到 {n} 筆）', 'status.stopping': '正在停止…',
      'status.error': '出錯：{msg}',
      'status.scanning': '掃描中…', 'status.scanDone': '掃描完', 'status.scanNone': '掃描完，未搵到訂單', 'status.scanErr': '掃描出錯',
      'stats.obtained': '已取得 {n} 筆訂單', 'stats.total': '合計 ¥{amt}',
      'col.createAt': '下單時間', 'col.orderSn': '訂單號', 'col.name': '商品名稱', 'col.spec': '規格',
      'col.price': '單價', 'col.qty': '數量', 'col.amount': '實付金額(元)', 'col.mall': '店鋪',
      'col.status': '狀態', 'col.type': '類型', 'col.buyUrl': '購買連結',
      'type.normal': '拼多多', 'type.fresh': '多多買菜',
      'summary.title': '拼多多買家訂單統計', 'summary.exportTime': '匯出時間', 'summary.orders': '訂單筆數',
      'summary.totalAmount': '合計金額(元)', 'summary.byStore': '按店鋪', 'summary.store': '店鋪',
      'summary.count': '筆數', 'summary.amount': '金額(元)', 'summary.byStatus': '按狀態', 'summary.status': '狀態',
      'summary.noStore': '(無店鋪)', 'summary.noStatus': '(無狀態)',
      'sheet.detail': '訂單明細', 'sheet.summary': '統計總覽',
      'export.none': '未有可匯出嘅訂單。', 'export.xlsxFail': '無法載入 Excel 元件，改為匯出 CSV。',
      'load.fail': '載入失敗', 'load.net': '無法連到 CDN',
      'menu.export': '匯出 Excel', 'menu.loadall': '自動載入全部訂單',
    },
    'zh-CN': {
      'app.title': '拼多多订单导出工具',
      'help.btn': '使用说明', 'close': '关闭', 'status.prefix': '状态：',
      'theme.btn': '暗黑模式', 'theme.toDark': '切换至暗黑模式', 'theme.toLight': '切换至浅色模式',
      'help.title': '怎么用？', 'help.ok': '知道了', 'help.close': '关闭/返回',
      'help.body': `<ol>
        <li>打开 <b>拼多多网页版</b> 并登录。</li>
        <li>进入「我的订单」→「<b>查看全部</b>」订单列表页面。</li>
        <li>选择<b>日期范围</b>（从／到）。</li>
        <li>（可选）点「选择导出字段」勾选想导出的项目。</li>
        <li>点「<b>搜索 &amp; 导出 Excel</b>」：会自动加载订单、按日期过滤、再导出 Excel（含统计总览）。</li>
        <li>加载期间可点橙色「<b>停止搜索</b>」随时停。</li>
      </ol><p>金额会自动由「分」换算为「元」；你选过的导出字段会记住。</p>`,
      'field.range': '日期范围', 'field.from': '从', 'field.to': '到',
      'field.cols': '选择导出字段（已选 {n} 项）▾', 'field.colsOpen': '收起', 'field.colsAll': '全选 / 全不选',
      'search': '搜索 & 导出 Excel', 'searchStop': '⏹ 停止搜索',
      'note': '点「搜索」后会自动加载订单，只导出所选日期范围内的订单。',
      'status.idle': '选择日期后点「搜索」',
      'status.searching': '搜索中… 正在加载订单', 'status.loading': '加载订单中…',
      'status.none': '该日期区间没有找到订单',
      'status.found': '找到 {n} 笔，导出中…', 'status.done': '完成，已导出 {n} 笔',
      'status.stopped': '已停止（已找到 {n} 笔）', 'status.stopping': '正在停止…',
      'status.error': '出错：{msg}',
      'status.scanning': '扫描中…', 'status.scanDone': '扫描完成', 'status.scanNone': '扫描完成，未找到订单', 'status.scanErr': '扫描出错',
      'stats.obtained': '已获取 {n} 笔订单', 'stats.total': '合计 ¥{amt}',
      'col.createAt': '下单时间', 'col.orderSn': '订单号', 'col.name': '商品名称', 'col.spec': '规格',
      'col.price': '单价', 'col.qty': '数量', 'col.amount': '实付金额(元)', 'col.mall': '店铺',
      'col.status': '状态', 'col.type': '类型', 'col.buyUrl': '购买链接',
      'type.normal': '拼多多', 'type.fresh': '多多买菜',
      'summary.title': '拼多多买家订单统计', 'summary.exportTime': '导出时间', 'summary.orders': '订单笔数',
      'summary.totalAmount': '合计金额(元)', 'summary.byStore': '按店铺', 'summary.store': '店铺',
      'summary.count': '笔数', 'summary.amount': '金额(元)', 'summary.byStatus': '按状态', 'summary.status': '状态',
      'summary.noStore': '(无店铺)', 'summary.noStatus': '(无状态)',
      'sheet.detail': '订单明细', 'sheet.summary': '统计总览',
      'export.none': '没有可导出的订单。', 'export.xlsxFail': '无法加载 Excel 组件，改为导出 CSV。',
      'load.fail': '加载失败', 'load.net': '无法连接到 CDN',
      'menu.export': '导出 Excel', 'menu.loadall': '自动加载全部订单',
    },
    'en': {
      'app.title': 'Pinduoduo Order Export',
      'help.btn': 'How to use', 'close': 'Close', 'status.prefix': 'Status: ',
      'theme.btn': 'Dark mode', 'theme.toDark': 'Switch to dark mode', 'theme.toLight': 'Switch to light mode',
      'help.title': 'How to use?', 'help.ok': 'Got it', 'help.close': 'Close/Back',
      'help.body': `<ol>
        <li>Open <b>Pinduoduo web</b> and sign in.</li>
        <li>Go to “My Orders” → “<b>View all</b>” order list page.</li>
        <li>Pick a <b>date range</b> (From / To).</li>
        <li>(Optional) Tap “Choose columns” to select which fields to export.</li>
        <li>Tap “<b>Search &amp; Export Excel</b>”: it loads orders, filters by date, then exports an Excel (with summary).</li>
        <li>During loading you can tap the orange “<b>Stop</b>” to cancel.</li>
      </ol><p>Amounts are automatically converted from cents to CNY; your chosen columns are remembered.</p>`,
      'field.range': 'Date range', 'field.from': 'From', 'field.to': 'To',
      'field.cols': 'Choose columns (selected {n}) ▾', 'field.colsOpen': 'Collapse', 'field.colsAll': 'Select all / None',
      'search': 'Search & Export Excel', 'searchStop': '⏹ Stop',
      'note': 'After “Search”, orders are loaded and only those in the selected date range are exported.',
      'status.idle': 'Pick a date range, then press “Search”',
      'status.searching': 'Searching… loading orders', 'status.loading': 'Loading orders…',
      'status.none': 'No orders in this date range',
      'status.found': 'Found {n}, exporting…', 'status.done': 'Done, exported {n}',
      'status.stopped': 'Stopped (found {n})', 'status.stopping': 'Stopping…',
      'status.error': 'Error: {msg}',
      'status.scanning': 'Scanning…', 'status.scanDone': 'Scan complete', 'status.scanNone': 'Scanned, no orders found', 'status.scanErr': 'Scan error',
      'stats.obtained': 'Captured {n} orders', 'stats.total': 'Total ¥{amt}',
      'col.createAt': 'Order time', 'col.orderSn': 'Order No.', 'col.name': 'Product', 'col.spec': 'Spec',
      'col.price': 'Unit price', 'col.qty': 'Qty', 'col.amount': 'Paid (CNY)', 'col.mall': 'Store',
      'col.status': 'Status', 'col.type': 'Type', 'col.buyUrl': 'Link',
      'type.normal': 'Pinduoduo', 'type.fresh': 'Duoduo Fresh',
      'summary.title': 'Pinduoduo Order Summary', 'summary.exportTime': 'Exported at', 'summary.orders': 'Orders',
      'summary.totalAmount': 'Total (CNY)', 'summary.byStore': 'By store', 'summary.store': 'Store',
      'summary.count': 'Count', 'summary.amount': 'Amount (CNY)', 'summary.byStatus': 'By status', 'summary.status': 'Status',
      'summary.noStore': '(no store)', 'summary.noStatus': '(no status)',
      'sheet.detail': 'Order Details', 'sheet.summary': 'Summary',
      'export.none': 'No orders to export.', 'export.xlsxFail': 'Cannot load Excel, exporting CSV instead.',
      'load.fail': 'Load failed', 'load.net': 'Cannot reach CDN',
      'menu.export': 'Export Excel', 'menu.loadall': 'Load all orders',
    },
  };

  const SUPPORTED = ['zh-HK', 'zh-CN', 'en'];
  function detectLocale() {
    const saved = store.get('lang', '');
    if (saved && SUPPORTED.includes(saved)) return saved;
    try {
      const l = (navigator.language || '').replace('_', '-');
      if (/^zh/i.test(l)) {
        if (/CN|Hans|Hans-|SG/i.test(l)) return 'zh-CN';
        return 'zh-HK';                          // 繁中（預設）
      }
      if (/^en/i.test(l)) return 'en';
    } catch (e) {}
    return 'zh-HK';
  }
  let locale = detectLocale();
  function t(key, params) {
    let s = (I18N[locale] && I18N[locale][key]) || I18N['zh-HK'][key] || I18N['en'][key] || key;
    if (params) for (const k in params) s = s.replace(new RegExp('\\{' + k + '\\}', 'g'), String(params[k]));
    return s;
  }
  function setLocale(l) { if (SUPPORTED.includes(l)) { locale = l; store.set('lang', l); } }

  // 匯出欄位定義（key）
  const COLUMNS = [
    { key: 'createAt' }, { key: 'orderSn' }, { key: 'name' }, { key: 'spec' }, { key: 'price' },
    { key: 'qty' }, { key: 'amount' }, { key: 'mall' }, { key: 'status' }, { key: 'type' }, { key: 'buyUrl' },
  ];
  const colLabel = (c) => t('col.' + c.key);
  const settings = {
    divide100: false,             // 若數量單位係「分」，開啟後全數除以 100
    selected: Object.fromEntries(COLUMNS.map((c) => [c.key, true])),
    dateFrom: '',                 // 'YYYY-MM-DD'
    dateTo: '',
    keyword: '',
  };

  /* ------------------------------------------------------------------ *
   *  載入上次嘅選項（只記「匯出欄位」，日期範圍唔記）
   * ------------------------------------------------------------------ */
  try {
    const sSel = store.get('selected', null);
    if (sSel && typeof sSel === 'object') for (const c of COLUMNS) if (typeof sSel[c.key] === 'boolean') settings.selected[c.key] = sSel[c.key];
  } catch (e) {}
  const saveSelected = () => store.set('selected', settings.selected);

  // 主題：記住用戶揀過（淺/深/系統預設）
  let themeMode = store.get('theme', '');   // '' = 跟系統, 'light', 'dark'
  function effectiveDark() {
    if (themeMode === 'dark') return true;
    if (themeMode === 'light') return false;
    try { return !!(window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches); } catch (e) { return false; }
  }

  /* ------------------------------------------------------------------ *
   *  工具函數
   * ------------------------------------------------------------------ */
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  function pad2(n) { return String(n).padStart(2, '0'); }

  function fmtTime(v) {
    if (v == null || v === '') return '';
    let d;
    if (typeof v === 'number') {
      d = new Date(v < 1e12 ? v * 1000 : v);
    } else {
      const n = Number(v);
      if (!Number.isNaN(n)) return fmtTime(n);
      d = new Date(v);
    }
    if (Number.isNaN(d.getTime())) return String(v);
    return [d.getFullYear(), pad2(d.getMonth() + 1), pad2(d.getDate())].join('-') + ' ' +
           [pad2(d.getHours()), pad2(d.getMinutes()), pad2(d.getSeconds())].join(':');
  }

  function toNum(v) {
    if (v == null) return NaN;
    if (typeof v === 'number') return v;
    const s = String(v).replace(/[^\d.\-]/g, '');
    const n = parseFloat(s);
    return Number.isNaN(n) ? NaN : n;
  }

  function fmtMoney(v, fallback) {
    const n = toNum(v);
    if (Number.isNaN(n)) return fallback == null ? '' : String(fallback);
    return (Math.round(n * 100) / 100).toFixed(2);
  }

  function txt(v) {
    if (v == null) return '';
    return String(v);
  }

  // 匯出用：一個欄位值 -> 顯示字串
  function colValue(rec, key) {
    switch (key) {
      case 'createAt': return rec.createAt;
      case 'orderSn':  return rec.orderSn;
      case 'name':     return rec.name;
      case 'spec':     return rec.spec;
      case 'qty':      return txt(rec.qty);
      case 'amount':   return rec.amount;         // 已格式化 2 位小數
      case 'price':    return rec.price;
      case 'mall':     return rec.mall;
      case 'status':   return rec.status;
      case 'type':     return rec.typeCode ? t('type.' + rec.typeCode) : '';
      case 'buyUrl':   return rec.buyUrl;
      default:         return '';
    }
  }

  /* ------------------------------------------------------------------ *
   *  訂單正規化（參照原 project 嘅欄位規則，並兼容多種欄位名）
   * ------------------------------------------------------------------ */
  function firstOf(obj, keys) {
    for (const k of keys) {
      if (obj && obj[k] != null && obj[k] !== '') return obj[k];
    }
    return null;
  }

  // 回傳值同埋「邊個 key 命中」，用嚟判斷單位
  function firstOfKey(obj, keys) {
    for (const k of keys) {
      if (obj && obj[k] != null && obj[k] !== '') return { v: obj[k], k };
    }
    return null;
  }

  // 金額單位：snake_case API 欄位 / displayAmount 一律係「分」，要除 100
  function scaleMoney(key, raw, fallback) {
    const n = toNum(raw);
    if (Number.isNaN(n)) return fallback == null ? '' : String(fallback);
    const cents = key.includes('_') || key === 'displayAmount' || key === 'display_amount';
    return fmtMoney((cents || settings.divide100) ? n / 100 : n);
  }

  function normalizeOrder(o) {
    const rec = {
      orderSn: '', createAt: '', name: '', spec: '', price: '',
      qty: '', amount: '', mall: '', status: '', typeCode: '', buyUrl: '',
    };
    try {
      rec.orderSn = getOrderSn(o);
      const type = o.type;                       // 1 = 一般訂單, 2 = 多多買菜
      const inner = Array.isArray(o.orders) && o.orders[0];
      rec.typeCode = (type === 2 || inner) ? 'fresh' : 'normal';

      // 搵商品：幾種常見位置（陣列取 [0]，單件物件直接用）
      let goods = null;
      const goodsCands = [o.orderGoods, o.order_goods, o.goods, o.goodsList, o.goods_list, o.items, o.itemList, o.orderItems, o.products];
      for (const c of goodsCands) {
        if (Array.isArray(c) && c[0]) { goods = c[0]; break; }
        if (c && typeof c === 'object') { goods = c; break; }
      }
      if (!goods && inner) {
        const ic = [inner.orderGoods, inner.order_goods, inner.goods, inner.orderItems, inner.products];
        for (const c of ic) {
          if (Array.isArray(c) && c[0]) { goods = c[0]; break; }
          if (c && typeof c === 'object') { goods = c; break; }
        }
      }

      // 金額與單價
      const amtKeys = ['order_amount', 'orderAmount', 'display_amount', 'displayAmount', 'pay_amount', 'payAmount', 'total_amount', 'totalAmount', 'goods_amount', 'goodsAmount', 'amount'];
      let amt = firstOfKey(o, amtKeys);
      if (!amt && goods) amt = firstOfKey(goods, ['goods_price', 'goodsPrice', 'price']);
      if (amt) rec.amount = scaleMoney(amt.k, amt.v);

      if (goods) {
        const pr = firstOfKey(goods, ['goods_price', 'goodsPrice', 'price', 'display_price', 'displayPrice', 'unit_price', 'unitPrice', 'sku_price']);
        if (pr) rec.price = scaleMoney(pr.k, pr.v);
        rec.name = txt(firstOf(goods, ['goods_name', 'goodsName', 'name', 'title']));
        rec.spec = txt(firstOf(goods, ['spec', 'goods_spec', 'goodsSpec', 'sku_name', 'sku']));
        rec.qty = toNum(firstOf(goods, ['goods_number', 'goodsNumber', 'quantity', 'num', 'count']));
        const gid = firstOf(goods, ['goods_id', 'goodsId']);
        if (gid) rec.buyUrl = 'https://mobile.pinduoduo.com/goods.html?goods_id=' + gid;
      }

      const t = firstOf(o, ['order_time', 'orderTime', 'create_time', 'createTime', 'createdAt']);
      rec.createAt = fmtTime(t != null ? t : (o.sortId ? Number(String(o.sortId).slice(0, 10)) : null));
      rec.mall = txt((o.mall && (o.mall.mall_name || o.mall.mallName || o.mall.name)) || firstOf(o, ['mall_name', 'mallName', 'shopName', 'merchantName']) || '');
      rec.status = txt(firstOf(o, ['order_status_prompt', 'orderStatusPrompt', 'order_status', 'orderStatus', 'statusText']) || '');
    } catch (e) { /* 忽略壞記錄 */ }
    return rec;
  }

  /* ------------------------------------------------------------------ *
   *  資料抽取：遞迴搵出「似訂單」嘅物件
   *  兼容多種訂單欄位名：orderSn / order_sn / orderId / order_id
   * ------------------------------------------------------------------ */
  function isOrderLike(o) {
    return !!o && typeof o === 'object' && (
      'orderSn' in o || 'order_sn' in o || 'orderId' in o || 'order_id' in o ||
      ('orderGoods' in o && 'orderAmount' in o)
    );
  }

  function getOrderSn(o) {
    if (o.orderSn != null && o.orderSn !== '') return txt(o.orderSn);
    if (o.order_sn != null && o.order_sn !== '') return txt(o.order_sn);
    if (o.orderId != null && o.orderId !== '') return txt(o.orderId);
    if (o.order_id != null && o.order_id !== '') return txt(o.order_id);
    if (o.orderNo != null && o.orderNo !== '') return txt(o.orderNo);
    return txt(o.order_no || '');
  }

  function findOrders(node, out, depth) {
    if (!node || depth > 8 || findBudget-- <= 0) return;
    if (Array.isArray(node)) {
      // 直接掃陣列元素（常見：list 就係訂單陣列）
      for (const item of node) {
        if (item && typeof item === 'object') {
          if (isOrderLike(item)) { out.push(item); }
          else findOrders(item, out, depth + 1);
        }
      }
    } else if (typeof node === 'object') {
      if (isOrderLike(node)) { out.push(node); return; }
      for (const k in node) { findOrders(node[k], out, depth + 1); }
    }
  }

  function addOrder(o, method) {
    if (!isOrderLike(o)) return false;
    const sn = getOrderSn(o);
    if (!sn || SEEN.has(o)) return false;
    SEEN.add(o);
    const rec = normalizeOrder(o);
    const prev = ORDER[sn];
    if (prev) {
      // 若之前只有部分資料，嘗試合併完善
      for (const k in rec) if (!prev[k] && rec[k]) prev[k] = rec[k];
      return false;
    }
    ORDER[sn] = rec;
    RAW[sn] = o;
    if (method) detectionMethod = method;
    return true;
  }

  // 記錄被攔截回應嘅結構（偵錯）
  const interceptSamples = [];
  function recordSample(url, text) {
    if (interceptSamples.length >= 8) return;
    let shape = '';
    try {
      const j = JSON.parse(text);
      shape = (Array.isArray(j) ? 'array[' + j.length + ']' :
               (j && typeof j === 'object' ? '{' + Object.keys(j).slice(0, 8).join(',') + '}' : typeof j));
    } catch (e) { shape = '(非JSON)'; }
    interceptSamples.push({ url, shape, len: (text || '').length, hasOrderSn: /\b(?:orderSn|order_sn|orderId)\b/.test(text || '') });
  }

  function handleJson(url, text) {
    interceptCount++;
    lastUrl = url;
    findBudget = 1e6;
    if (!text || text.length < 2) { dbg('resp (empty body)', url); return; }
    // 只處理 JSON 內容
    const t = text.trim();
    if (!(t.startsWith('{') || t.startsWith('['))) { dbg('resp (non-json)', url, t.slice(0, 60)); return; }
    if (text.length > 8 * 1024 * 1024) return;     // 過大跳過
    let json;
    try { json = JSON.parse(text); } catch (e) { dbg('resp (json parse fail)', url); return; }
    recordSample(url, text);
    const found = [];
    findOrders(json, found, 0);
    dbg('resp #' + interceptCount, url, 'ordersFound=' + found.length);
    let any = false;
    for (const o of found) if (addOrder(o, 'XHR/API 攔截')) any = true;
    if (any && windowOn()) scheduleRender();
  }

  /* ------------------------------------------------------------------ *
   *  後備：掃描 React fiber（等同原 project 讀取內部 state）
   *  非阻塞寫法：分批「讓出主執行緒」，唔會卡死頁面。
   * ------------------------------------------------------------------ */
  const yieldUI = () => new Promise((r) => setTimeout(r, 0));
  let scanning = false;

  async function collectFromFiber() {
    const out = [];
    const seen = new Set();
    const roots = new Set();
    // 只掃需要嘅元素，搵到唔同嘅最高層 fiber root 就停，唔逐個元素起步
    const els = document.querySelectorAll('div,[data-reactroot]');
    for (let i = 0; i < els.length && roots.size < 6 && i < 600; i++) {
      const el = els[i];
      let f = null;
      for (const k of Object.keys(el)) {
        if (k.startsWith('__reactFiber$') || k.startsWith('__reactInternalInstance$')) { f = el[k]; break; }
      }
      if (!f || typeof f !== 'object') continue;
      let top = f, g = 0;
      while (top && top.return && g++ < 500) top = top.return;  // 向上爬去 root
      if (top) roots.add(top);
      if ((i & 0x1f) === 0) await yieldUI();
    }

    const stack = [...roots];
    let guard = 0;
    let lastYield = Date.now();
    const startAt = Date.now();
    while (stack.length && guard++ < 30000) {
      if (cancelled) break;                                 // 用戶停止
      const f = stack.pop();
      if (!f || typeof f !== 'object' || seen.has(f)) continue;
      seen.add(f);
      findBudget = 9000;                               // 每次 findOrders 嘅小預算（防單次卡死）
      let hs = f.memoizedState, g = 0;                 // hooks 鏈
      while (hs && g++ < 200) {
        if (hs.memoizedState) findOrders(hs.memoizedState, out, 0);
        hs = hs.next;
      }
      if (f.memoizedProps) findOrders(f.memoizedProps, out, 0);   // props（如 itemsStore）
      if (f.child) stack.push(f.child);
      if (f.sibling) stack.push(f.sibling);
      // 按時間讓出主執行緒：每次最多跑 ~12ms 就交返畀瀏覽器
      if (Date.now() - lastYield > 12) { await yieldUI(); lastYield = Date.now(); }
      // 硬上限：最多行 2 秒，避免「無止境掃描」
      if (Date.now() - startAt > 2000) break;
    }
    return out;
  }

  async function scanPage() {
    if (scanning) return;
    scanning = true;
    setStatus(t('status.scanning'));
    renderStats();
    try {
      const found = await collectFromFiber();
      let any = false;
      for (const o of found) if (addOrder(o, '頁面掃描(React)')) any = true;
      dbg('scanPage found=' + found.length);
      renderStats();
      setStatus(any ? t('status.scanDone') : t('status.scanNone'));
      return any;
    } catch (e) {
      dbg('scanPage error', e);
      setStatus(t('status.scanErr'));
      return false;
    } finally {
      scanning = false;
    }
  }

  /* ------------------------------------------------------------------ *
   *  網路攔截：hook fetch + XMLHttpRequest
   * ------------------------------------------------------------------ */
  function hookNetwork() {
    const W = unsafeWindow;
    if (W && typeof W.fetch === 'function' && !W.fetch.__pddHooked) {
      const origFetch = W.fetch;
      const hooked = function (input, init) {
        const url = typeof input === 'string' ? input : (input && input.url) || '';
        return origFetch.apply(this, arguments).then((resp) => {
          try {
            if (resp && typeof resp.clone === 'function') {
              const clone = resp.clone();
              clone.text().then((t) => handleJson(url, t)).catch(() => {});
            }
          } catch (e) {}
          return resp;
        });
      };
      hooked.__pddHooked = true;
      W.fetch = hooked;
    }

    const XHR = W && W.XMLHttpRequest;
    if (XHR && XHR.prototype && !XHR.prototype.__pddHooked) {
      const origOpen = XHR.prototype.open;
      const origSend = XHR.prototype.send;
      XHR.prototype.open = function (method, url) {
        this.__pddUrl = url;
        return origOpen.apply(this, arguments);
      };
      XHR.prototype.send = function () {
        try {
          const self = this;
          this.addEventListener('load', function () {
            try {
              let body = self.responseText;
              if (!body && self.response != null) {
                body = typeof self.response === 'string' ? self.response
                     : (() => { try { return JSON.stringify(self.response); } catch (e) { return ''; } })();
              }
              handleJson(self.__pddUrl || '', body || '');
            } catch (e) {}
          });
        } catch (e) {}
        return origSend.apply(this, arguments);
      };
      XHR.prototype.__pddHooked = true;
    }
    dbg('hook installed. fetch=' + (W && typeof W.fetch === 'function') + ' xhr=' + !!XHR);
  }

  function oldestDate() {
    let min = null;
    for (const r of Object.values(ORDER)) {
      if (r.createAt) { const d = r.createAt.slice(0, 10); if (!min || d < min) min = d; }
    }
    return min;
  }

  /* ------------------------------------------------------------------ *
   *  自動載入：往下滾動，直到載晒「指定日期範圍」嘅訂單（新→舊所以唔使載晒全部）
   * ------------------------------------------------------------------ */
  async function autoLoadAll() {
    if (autoLoading) return;
    autoLoading = true;
    setStatus(t('status.loading'));
    let prev = Object.keys(ORDER).length;
    let stable = 0;
    let passedStart = 0;
    const startAt = Date.now();
    for (let i = 0; i < 260; i++) {
      window.scrollTo(0, document.body.scrollHeight);
      await sleep(600);
      const n = Object.keys(ORDER).length;
      if (n === prev) { stable++; } else { stable = 0; prev = n; }
      renderStats();
      if (cancelled) break;                                    // 用戶停止
      // 若設咗「從」日期，而且已載到「比從日期仲舊」嘅訂單 => 範圍已覆蓋，可以停
      if (settings.dateFrom) {
        const od = oldestDate();
        if (od && od <= settings.dateFrom) passedStart++; else passedStart = 0;
        if (passedStart >= 2) break;
      }
      if (stable >= 5) break;                                  // 到底
      if (Date.now() - startAt > 45000) break;                 // 硬上限
    }
    autoLoading = false;
    renderStats();
  }

  /* ------------------------------------------------------------------ *
   *  篩選過濾
   * ------------------------------------------------------------------ */
  function inRange(rec) {
    const from = settings.dateFrom, to = settings.dateTo;
    if (from || to) {
      const d = rec.createAt ? new Date(rec.createAt) : null;
      if (d && !Number.isNaN(d.getTime())) {
        const day = [d.getFullYear(), pad2(d.getMonth() + 1), pad2(d.getDate())].join('-');
        if (from && day < from) return false;
        if (to && day > to) return false;
      }
    }
    if (settings.keyword) {
      const kw = settings.keyword.toLowerCase();
      const hay = (rec.name + ' ' + rec.spec + ' ' + rec.mall + ' ' + rec.orderSn).toLowerCase();
      if (!hay.includes(kw)) return false;
    }
    return true;
  }

  function filtered() {
    return Object.values(ORDER).filter(inRange);
  }

  /* ------------------------------------------------------------------ *
   *  UI
   * ------------------------------------------------------------------ */
  let els = null;

  function windowOn() { return typeof window !== 'undefined' && !!document.body; }

  function registerUI() {
    if (els) return;
    const host = document.createElement('div');
    host.id = NS;
    host.innerHTML = PANEL_HTML;
    document.body.appendChild(host);

    els = {
      root: host,
      fab: host.querySelector('.pdd-fab'),
      panel: host.querySelector('.pdd-panel'),
      stats: host.querySelector('.pdd-stats'),
      status: host.querySelector('.pdd-status-val'),
      toggle: host.querySelector('.pdd-toggle'),
      search: host.querySelector('.pdd-search'),
      from: host.querySelector('.pdd-from'),
      to: host.querySelector('.pdd-to'),
      colsToggle: host.querySelector('.pdd-cols-toggle'),
      cols: host.querySelector('.pdd-cols'),
      colsWrap: host.querySelector('.pdd-cols-wrap'),
      colsAll: host.querySelector('.pdd-cols-all'),
      theme: host.querySelector('.pdd-theme'),
      helpbtn: host.querySelector('.pdd-helpbtn'),
      help: host.querySelector('.pdd-help'),
      helpClose: host.querySelector('.pdd-help-close'),
      helpOk: host.querySelector('.pdd-help-ok'),
      helpBody: host.querySelector('.pdd-help-body'),
      lang: host.querySelector('.pdd-lang'),
    };

    els.fab.addEventListener('click', () => els.panel.classList.toggle('pdd-open'));
    els.toggle.addEventListener('click', () => els.panel.classList.remove('pdd-open'));
    els.search.addEventListener('click', () => { if (searching) cancelSearch(); else onSearchExport(); });
    els.from.addEventListener('change', () => { settings.dateFrom = els.from.value; renderStats(); });
    els.to.addEventListener('change', () => { settings.dateTo = els.to.value; renderStats(); });
    els.colsToggle.addEventListener('click', () => {
      const open = els.colsWrap.classList.toggle('pdd-open');
      els.colsToggle.textContent = (open ? t('field.colsOpen') + ' ' : '') + t('field.cols', { n: selectedCount() });
    });
    els.colsAll.addEventListener('click', () => {
      const all = selectedCount() < COLUMNS.length;   // 未全選 => 全選；已全選 => 全不選
      for (const c of COLUMNS) settings.selected[c.key] = all;
      buildColumnCheckboxes();
      updateColsLabel();
      saveSelected();
    });
    els.theme.addEventListener('click', () => {
      els.root.setAttribute('data-theme-switching', '');
      themeMode = effectiveDark() ? 'light' : 'dark';
      store.set('theme', themeMode);
      applyTheme();
      setTimeout(() => els.root.removeAttribute('data-theme-switching'), 320);
    });
    els.lang.addEventListener('change', () => { setLocale(els.lang.value); applyI18n(); });
    const openHelp = () => els.help.classList.add('pdd-open');
    const closeHelp = () => els.help.classList.remove('pdd-open');
    els.helpbtn.addEventListener('click', openHelp);
    els.helpClose.addEventListener('click', closeHelp);
    els.helpOk.addEventListener('click', closeHelp);

    // 日期範圍唔記，每次重新開始（留空）
    els.from.value = '';
    els.to.value = '';

    applyI18n();
    setStatus(t('status.idle'));
  }

  function selectedCount() { return COLUMNS.filter((c) => settings.selected[c.key]).length; }

  function updateColsLabel() {
    if (els) els.colsToggle.textContent = t('field.cols', { n: selectedCount() });
  }

  function buildColumnCheckboxes() {
    if (!els) return;
    const box = els.cols;
    // 保留「全選 / 全不選」按鈕，只重建 checkbox
    box.querySelectorAll('.pdd-col').forEach((n) => n.remove());
    for (const c of COLUMNS) {
      const label = document.createElement('label');
      label.className = 'pdd-col';
      const inp = document.createElement('input');
      inp.type = 'checkbox';
      inp.checked = settings.selected[c.key];
      inp.addEventListener('change', () => { settings.selected[c.key] = inp.checked; updateColsLabel(); saveSelected(); });
      label.appendChild(inp);
      const span = document.createElement('span');
      span.textContent = colLabel(c);
      label.appendChild(span);
      box.appendChild(label);
    }
  }

  function setStatus(s) { if (els) els.status.textContent = s; }

  function applyTheme() {
    if (!els) return;
    const dark = effectiveDark();
    els.root.dataset.theme = dark ? 'dark' : 'light';
    els.theme.textContent = dark ? '☀️' : '🌙';
    els.theme.title = dark ? t('theme.toLight') : t('theme.toDark');
  }

  // 套用當前語言：更新所有標記咗 data-i18n / data-i18n-title 嘅文字 + 動態文字
  function applyI18n() {
    if (!els) return;
    els.root.querySelectorAll('[data-i18n]').forEach((el) => { el.textContent = t(el.dataset.i18n); });
    els.root.querySelectorAll('[data-i18n-title]').forEach((el) => { el.title = t(el.dataset.i18nTitle); });
    if (els.helpBody) els.helpBody.innerHTML = t('help.body');
    els.lang.value = locale;
    buildColumnCheckboxes();
    updateColsLabel();
    renderStats();
    applyTheme();
  }

  function renderStats() {
    if (!els) return;
    const all = Object.values(ORDER);
    const totalAmount = all.reduce((s, r) => s + (Number.isNaN(toNum(r.amount)) ? 0 : toNum(r.amount)), 0);
    els.stats.textContent = t('stats.obtained', { n: all.length }) + '  |  ' + t('stats.total', { amt: fmtMoney(totalAmount) });
  }

  function scheduleRender() {
    if (scheduleRender._t) return;
    scheduleRender._t = setTimeout(() => { scheduleRender._t = null; renderStats(); }, 250);
  }

  // 一鍵：載入全部 -> 過濾日期範圍 -> 匯出 Excel（搜尋期間可停止）
  let searching = false;
  let cancelled = false;
  async function onSearchExport() {
    if (searching) return;
    searching = true; cancelled = false;
    setButton(t('searchStop'), true);
    setStatus(t('status.searching'));
    try {
      await autoLoadAll();                 // 向下滾動，網路攔截自動累積
      if (cancelled) { renderStats(); return; }
      if (Object.keys(ORDER).length === 0) {
        await scanPage();                  // 只有網路完全捉唔到先掃描（有時間上限）
      }
      if (cancelled) { renderStats(); return; }
      renderStats();
      const rows = currentRows();          // 依日期範圍過濾
      if (!rows.length) { setStatus(t('status.none')); return; }
      setStatus(t('status.found', { n: rows.length }));
      await exportXlsx();
      if (!cancelled) setStatus(t('status.done', { n: rows.length }));
    } catch (e) {
      setStatus(t('status.error', { msg: ((e && e.message) || e) }));
    } finally {
      if (cancelled) setStatus(t('status.stopped', { n: Object.keys(ORDER).length }));
      searching = false;
      setButton(t('search'), false);
    }
  }

  function cancelSearch() { if (searching) { cancelled = true; setStatus(t('status.stopping')); } }

  function setButton(text, busy) {
    if (!els) return;
    els.search.textContent = text;
    els.search.disabled = false;
    els.search.classList.toggle('pdd-busy', !!busy);
  }

  /* ------------------------------------------------------------------ *
   *  匯出
   * ------------------------------------------------------------------ */
  function download(filename, content, mime) {
    const blob = content instanceof Blob ? content : new Blob([content], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 1000);
  }

  function buildAoa(rows) {
    let shown = COLUMNS.filter((c) => settings.selected[c.key]);
    if (!shown.length) shown = COLUMNS;             // 冇揀任何欄位時，用晒全部，避免空表
    const header = shown.map((c) => colLabel(c));
    const body = rows.map((r) => shown.map((c) => colValue(r, c.key)));
    return { header, body, shown };
  }

  function currentRows() {
    return filtered().sort((a, b) => {
      const ta = a.createAt ? +new Date(a.createAt) : 0;
      const tb = b.createAt ? +new Date(b.createAt) : 0;
      return tb - ta;
    });
  }

  function exportCsv() {
    const rows = currentRows();
    if (!rows.length) { alert(t('export.none')); return; }
    const { header, body } = buildAoa(rows);
    let csv = '\uFEFF'; // BOM，Excel 正確顯示中文
    csv += header.map(csvCell).join(',') + '\r\n';
    for (const r of body) csv += r.map(csvCell).join(',') + '\r\n';
    download('pdd-orders-' + stamp() + '.csv', csv, 'text/csv;charset=utf-8');
  }

  function csvCell(v) {
    const s = txt(v);
    if (/[",\r\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
    return s;
  }

  function stamp() {
    const d = new Date();
    return [d.getFullYear(), pad2(d.getMonth() + 1), pad2(d.getDate())].join('') + '-' +
           [pad2(d.getHours()), pad2(d.getMinutes()), pad2(d.getSeconds())].join('');
  }

  async function exportXlsx() {
    const rows = currentRows();
    if (!rows.length) { alert(t('export.none')); return; }
    let XLSX;
    try {
      XLSX = await loadXLSX();
    } catch (e) {
      alert(t('export.xlsxFail') + '\n' + e.message);
      exportCsv();
      return;
    }

    const { header, body } = buildAoa(rows);
    const wsData = [header, ...body];
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    ws['!cols'] = header.map((h, i) => ({ wch: colWidth(header[i], body, i) }));

    // 統計總覽分頁
    const summary = buildSummary(rows);
    const ws2 = XLSX.utils.aoa_to_sheet(summary);
    ws2['!cols'] = [{ wch: 26 }, { wch: 16 }, { wch: 16 }];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, t('sheet.detail'));
    XLSX.utils.book_append_sheet(wb, ws2, t('sheet.summary'));

    const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    download('pdd-orders-' + stamp() + '.xlsx', new Blob([out], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  }

  function colWidth(h, body, i) {
    let w = h.length; for (const r of body) w = Math.max(w, txt(r[i]).length);
    return Math.min(Math.max(w + 2, 8), 48);
  }

  function buildSummary(rows) {
    const totalAmount = rows.reduce((s, r) => s + (Number.isNaN(toNum(r.amount)) ? 0 : toNum(r.amount)), 0);
    const lines = [
      [t('summary.title'), '', ''],
      [t('summary.exportTime'), fmtTime(Date.now()), ''],      [t('summary.orders'), rows.length, ''],
      [t('summary.totalAmount'), fmtMoney(totalAmount), ''],
      ['', '', ''],
      [t('summary.byStore'), '', ''],
      [t('summary.store'), t('summary.count'), t('summary.amount')],
    ];
    const byMall = {};
    const byStatus = {};
    for (const r of rows) {
      const m = r.mall || t('summary.noStore');
      byMall[m] = byMall[m] || { n: 0, amt: 0 };
      byMall[m].n++; byMall[m].amt += (Number.isNaN(toNum(r.amount)) ? 0 : toNum(r.amount));
      const s = r.status || t('summary.noStatus');
      byStatus[s] = (byStatus[s] || 0) + 1;
    }
    for (const m in byMall) lines.push([m, byMall[m].n, fmtMoney(byMall[m].amt)]);
    lines.push(['', '', '']);
    lines.push([t('summary.byStatus'), '', '']);
    lines.push([t('summary.status'), t('summary.count'), '']);
    for (const s in byStatus) lines.push([s, byStatus[s], '']);
    return lines;
  }

  let _xlsxPromise = null;
  function loadXLSX() {
    if (unsafeWindow.XLSX) return Promise.resolve(unsafeWindow.XLSX);
    if (_xlsxPromise) return _xlsxPromise;
    _xlsxPromise = new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
      s.onload = () => (unsafeWindow.XLSX ? resolve(unsafeWindow.XLSX) : reject(new Error(t('load.fail'))));
      s.onerror = () => reject(new Error(t('load.net')));
      document.head.appendChild(s);
    });
    return _xlsxPromise;
  }

  /* ------------------------------------------------------------------ *
   *  UI template
   * ------------------------------------------------------------------ */
  const PANEL_HTML = `
  <style>
    #${NS}{all:initial;font-family:-apple-system,"PingFang HK","Microsoft JhengHei",sans-serif;color:var(--pdd-text);}
    #${NS} button{font-family:inherit;}
    #${NS} *{box-sizing:border-box;}

    /* 主題變數（淺色預設） */
    #${NS}{
      --pdd-bg:#ffffff; --pdd-bg2:#f7f8f9; --pdd-border:#e6e6e6; --pdd-text:#222; --pdd-sub:#666; --pdd-muted:#999;
      --pdd-accent:#e02e24; --pdd-green:#217346; --pdd-green-hover:#1a5c37; --pdd-input-bg:#fff; --pdd-input-border:#ddd;
      --pdd-tgl-bg:#f2f3f4; --pdd-check:#217346; --pdd-shadow:rgba(0,0,0,.25); --pdd-fab-shadow:rgba(0,0,0,.28);
    }
    #${NS}[data-theme="dark"]{
      --pdd-bg:#1d2024; --pdd-bg2:#262a30; --pdd-border:#3a3f45; --pdd-text:#e8e9ea; --pdd-sub:#b3b6ba; --pdd-muted:#8a8d90;
      --pdd-accent:#ff5148; --pdd-green:#2fa36a; --pdd-green-hover:#39b577; --pdd-input-bg:#2a2e33; --pdd-input-border:#454b52;
      --pdd-tgl-bg:#30343a; --pdd-check:#2fa36a; --pdd-shadow:rgba(0,0,0,.55); --pdd-fab-shadow:rgba(0,0,0,.5);
    }
    #${NS}[data-theme-switching] *{transition:none !important;}

    .pdd-fab{position:fixed;right:18px;bottom:18px;z-index:2147483647;background:var(--pdd-accent);color:#fff;border:none;
      border-radius:999px;padding:12px 18px;font-size:15px;font-weight:600;cursor:pointer;box-shadow:0 4px 16px var(--pdd-fab-shadow);
      transition:transform .16s cubic-bezier(.22,1,.36,1),box-shadow .16s ease,background-color .16s ease;}
    .pdd-panel{position:fixed;right:18px;bottom:64px;z-index:2147483646;width:360px;background:var(--pdd-bg);
      border:1px solid var(--pdd-border);border-radius:14px;overflow:hidden;box-shadow:0 10px 40px var(--pdd-shadow);
      opacity:0;transform:scale(.94) translateY(14px);transform-origin:right bottom;visibility:hidden;pointer-events:none;
      transition:transform .28s cubic-bezier(.22,1,.36,1),opacity .26s ease,visibility 0s linear .26s;}
    .pdd-panel.pdd-open{opacity:1;transform:none;visibility:visible;pointer-events:auto;transition:transform .28s cubic-bezier(.22,1,.36,1),opacity .26s ease,visibility 0s;}
    .pdd-head{display:flex;align-items:center;justify-content:space-between;padding:12px 16px;border-bottom:1px solid var(--pdd-border);
      background:var(--pdd-bg2);border-radius:14px 14px 0 0;transition:background-color .2s ease,border-color .2s ease;}
    .pdd-head b{font-size:15px;color:var(--pdd-text);}
    .pdd-head-left{display:flex;align-items:center;gap:7px;}
    .pdd-helpbtn{background:none;border:1px solid var(--pdd-border);color:var(--pdd-sub);width:20px;height:20px;border-radius:50%;
      font-size:12px;line-height:1;cursor:pointer;display:inline-flex;align-items:center;justify-content:center;flex:none;
      transition:color .16s ease,background-color .16s ease,border-color .16s ease,transform .12s ease;}
    .pdd-helpbtn:hover{background:var(--pdd-tgl-bg);color:var(--pdd-green);border-color:var(--pdd-green);}
    .pdd-helpbtn:active{transform:scale(.9);}
    .pdd-head-actions{display:flex;gap:4px;align-items:center;}
    .pdd-hbtn{background:none;border:none;color:var(--pdd-sub);font-size:16px;cursor:pointer;line-height:1;padding:4px 6px;border-radius:6px;
      transition:color .16s ease,background-color .16s ease,transform .12s ease;}
    .pdd-hbtn:hover{background:var(--pdd-tgl-bg);color:var(--pdd-text);}
    .pdd-hbtn:active{transform:scale(.88);}
    .pdd-lang{background:var(--pdd-tgl-bg);color:var(--pdd-sub);border:1px solid var(--pdd-border);border-radius:6px;font-size:11.5px;
      padding:3px 4px;cursor:pointer;transition:color .16s ease,border-color .16s ease;}
    .pdd-lang:hover{color:var(--pdd-text);border-color:var(--pdd-green);}
    .pdd-lang option{background:var(--pdd-bg);color:var(--pdd-text);}
    .pdd-body{padding:16px;}
    .pdd-status{margin:0 0 6px;font-size:12px;color:var(--pdd-sub);min-height:16px;}
    .pdd-stats{font-weight:600;color:var(--pdd-accent);font-size:14px;margin-bottom:14px;}
    .pdd-sec{font-size:12px;color:var(--pdd-muted);margin:0 0 6px;}
    .pdd-field{display:flex;gap:8px;align-items:center;margin-bottom:16px;}
    .pdd-field label{font-size:13px;color:var(--pdd-text);white-space:nowrap;}
    .pdd-field input[type=date]{flex:1;min-width:0;background:var(--pdd-input-bg);color:var(--pdd-text);
      border:1px solid var(--pdd-input-border);border-radius:7px;padding:8px;font-size:13px;
      transition:border-color .16s ease,box-shadow .16s ease;}
    .pdd-field input[type=date]:focus{outline:none;border-color:var(--pdd-green);box-shadow:0 0 0 3px rgba(47,163,106,.18);}
    .pdd-search{width:100%;border:none;border-radius:9px;padding:12px;font-size:15px;font-weight:600;color:#fff;
      background:var(--pdd-green);cursor:pointer;margin-top:2px;
      transition:transform .14s cubic-bezier(.22,1,.36,1),background-color .16s ease,box-shadow .16s ease;}
    .pdd-search:hover{background:var(--pdd-green-hover);}
    .pdd-search:active{transform:scale(.98);}
    .pdd-search:disabled{opacity:.55;cursor:default;}
    .pdd-search.pdd-busy{background:#c2701b;}
    .pdd-search.pdd-busy:hover{background:#a85c12;}
    .pdd-cols-toggle{width:100%;background:var(--pdd-tgl-bg);border:1px solid var(--pdd-border);border-radius:8px;padding:8px 12px;
      font-size:12.5px;color:var(--pdd-sub);cursor:pointer;margin:0 0 12px;text-align:left;
      transition:background-color .16s ease,color .16s ease,border-color .16s ease;}
    .pdd-cols-toggle:hover{background:var(--pdd-bg2);color:var(--pdd-text);}
    .pdd-cols-wrap{display:grid;grid-template-rows:0fr;transition:grid-template-rows .26s cubic-bezier(.22,1,.36,1);}
    .pdd-cols-wrap.pdd-open{grid-template-rows:1fr;}
    .pdd-cols{overflow:hidden;min-height:0;display:flex;flex-wrap:wrap;gap:6px 14px;padding:2px 4px 12px;
      opacity:0;transform:translateY(-6px);transition:opacity .22s ease,transform .22s ease;}
    .pdd-cols-wrap.pdd-open .pdd-cols{opacity:1;transform:none;}
    .pdd-col{display:flex;align-items:center;gap:5px;font-size:12.5px;color:var(--pdd-text);cursor:pointer;}
    .pdd-col input{accent-color:var(--pdd-check);}
    .pdd-col span{transition:color .16s ease;}
    .pdd-col:hover span{color:var(--pdd-green);}
    .pdd-cols-all{width:100%;font-size:12px;color:var(--pdd-green);background:none;border:none;cursor:pointer;padding:0 0 4px;text-align:left;
      transition:opacity .16s ease;}
    .pdd-cols-all:hover{opacity:.75;}
    .pdd-note{font-size:11px;color:var(--pdd-muted);margin-top:10px;line-height:1.5;}
    .pdd-help{position:absolute;inset:0;z-index:6;background:rgba(0,0,0,.42);display:flex;align-items:center;justify-content:center;
      opacity:0;visibility:hidden;pointer-events:none;border-radius:14px;transition:opacity .22s ease,visibility 0s linear .22s;}
    .pdd-help.pdd-open{opacity:1;visibility:visible;pointer-events:auto;transition:opacity .22s ease,visibility 0s;}
    .pdd-help-card{width:88%;max-height:84%;overflow:auto;background:var(--pdd-bg);border:1px solid var(--pdd-border);border-radius:12px;
      box-shadow:0 12px 44px var(--pdd-shadow);transform:scale(.94);opacity:0;transition:transform .26s cubic-bezier(.22,1,.36,1),opacity .22s ease;}
    .pdd-help.pdd-open .pdd-help-card{transform:none;opacity:1;}
    .pdd-help-head{display:flex;align-items:center;justify-content:space-between;padding:12px 16px;border-bottom:1px solid var(--pdd-border);
      background:var(--pdd-bg2);border-radius:12px 12px 0 0;}
    .pdd-help-head b{font-size:14px;color:var(--pdd-text);}
    .pdd-help-body{padding:14px 18px;font-size:13px;color:var(--pdd-sub);line-height:1.7;}
    .pdd-help-body ol{margin:0;padding-left:20px;}
    .pdd-help-body li{margin:0 0 9px;}
    .pdd-help-body p{margin:10px 0 0;color:var(--pdd-muted);font-size:12px;}
    .pdd-help-ok{display:block;width:calc(100% - 36px);margin:0 18px 16px;border:none;border-radius:8px;padding:10px;font-size:14px;font-weight:600;
      color:#fff;background:var(--pdd-green);cursor:pointer;transition:transform .14s cubic-bezier(.22,1,.36,1),background-color .16s ease;}
    .pdd-help-ok:hover{background:var(--pdd-green-hover);}
    .pdd-help-ok:active{transform:scale(.98);}
    @media (hover:hover) and (pointer:fine){
      .pdd-fab:hover{transform:translateY(-2px) scale(1.04);box-shadow:0 6px 22px var(--pdd-fab-shadow);}
      .pdd-fab:active{transform:scale(.95);}
    }
  </style>
  <button class="pdd-fab" data-i18n-title="app.title">訂單</button>
  <div class="pdd-panel">
    <div class="pdd-head">
      <span class="pdd-head-left">
        <b data-i18n="app.title">拼多多訂單導出工具</b>
        <button class="pdd-helpbtn" data-i18n-title="help.btn">?</button>
      </span>
      <span class="pdd-head-actions">
        <select class="pdd-lang" title="語言 / Language">
          <option value="zh-HK">繁體</option>
          <option value="zh-CN">简体</option>
          <option value="en">EN</option>
        </select>
        <button class="pdd-hbtn pdd-theme" title="暗黑模式">🌙</button>
        <button class="pdd-hbtn pdd-toggle" data-i18n-title="close">×</button>
      </span>
    </div>
    <div class="pdd-body">
      <p class="pdd-status"><span data-i18n="status.prefix">狀態：</span><span class="pdd-status-val" data-i18n="status.idle">揀日期後撳「搜尋」</span></p>
      <div class="pdd-stats">已取得 0 筆訂單</div>

      <div class="pdd-sec" data-i18n="field.range">日期範圍</div>
      <div class="pdd-field">
        <label data-i18n="field.from">從</label><input type="date" class="pdd-from">
        <label data-i18n="field.to">到</label><input type="date" class="pdd-to">
      </div>

      <button class="pdd-cols-toggle">選擇匯出欄位（已選 11 項）▾</button>
      <div class="pdd-cols-wrap">
        <div class="pdd-cols">
          <button class="pdd-cols-all" data-i18n="field.colsAll">全選 / 全不選</button>
        </div>
      </div>

      <button class="pdd-search" data-i18n="search">搜尋 & 匯出 Excel</button>
      <p class="pdd-note" data-i18n="note">撳「搜尋」後會自動載入訂單，只匯出所揀日期範圍內嘅訂單。</p>
    </div>

    <div class="pdd-help">
      <div class="pdd-help-card">
        <div class="pdd-help-head"><b data-i18n="help.title">點樣用？</b><button class="pdd-hbtn pdd-help-close" data-i18n-title="help.close">×</button></div>
        <div class="pdd-help-body"></div>
        <button class="pdd-help-ok" data-i18n="help.ok">知道了</button>
      </div>
    </div>
  </div>`;

  /* ------------------------------------------------------------------ *
   *  啟動
   * ------------------------------------------------------------------ */
  function boot() {
    hookNetwork();
    if (document.body) registerUI();
    else document.addEventListener('DOMContentLoaded', registerUI);

    // 唔同頁面（SPA 跳轉）都要確保 UI 存在
    const mo = new MutationObserver(() => { if (document.body && !els) registerUI(); });
    if (document.body) mo.observe(document.body, { childList: true, subtree: false });

    // 讓已存在嘅訂單資料（若頁面先前已載入）有機會被捕獲
    setInterval(() => { if (windowOn() && !els) registerUI(); }, 1500);
  }

  // menu 指令
  try {
    GM_registerMenuCommand(t('menu.export'), () => { if (windowOn()) exportXlsx(); });
    GM_registerMenuCommand(t('menu.loadall'), () => { if (windowOn()) autoLoadAll(); });
  } catch (e) {}

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', boot);
  } else {
    boot();
  }

  // 對外暴露（方便 debug）
  unsafeWindow.__pddOrderExport = {
    ORDER, RAW, filtered, settings,
    exportXlsx, exportCsv, autoLoadAll, scanPage, cancelSearch,
    debug: (on = true) => { unsafeWindow.__pddDebug = !!on; },
    theme: () => ({ mode: themeMode, dark: effectiveDark() }),
    setTheme: (m) => { themeMode = m || ''; store.set('theme', themeMode); if (els) applyTheme(); },
    i18n: { t, locale: () => locale, setLocale, SUPPORTED },
    setLocale: (l) => { setLocale(l); if (els) applyI18n(); },
    stats: () => ({ count: Object.keys(ORDER).length, interceptCount, lastUrl, method: detectionMethod }),
    samples: () => interceptSamples.slice(),
    // 顯示已捕捉訂單嘅「真實欄位結構」，方便針對實際格式修正
    dump: () => {
      const first = Object.values(RAW)[0] || Object.values(ORDER)[0];
      if (!first) return { error: '未有已捕捉嘅訂單' };
      const shallow = {};
      for (const k in first) {
        const v = first[k];
        shallow[k] = Array.isArray(v) ? ('array[' + v.length + ']')
                   : (v && typeof v === 'object' ? ('{' + Object.keys(v).slice(0, 6).join(',') + '}') : v);
      }
      return { keys: Object.keys(first), sample: shallow };
    },
    dumpAll: () => Object.values(RAW).slice(0, 3).map((o) => {
      const s = {};
      for (const k in o) {
        const v = o[k];
        s[k] = Array.isArray(v) ? ('array[' + v.length + ']') : (v && typeof v === 'object' ? '{obj}' : v);
      }
      return s;
    }),
  };
})();
