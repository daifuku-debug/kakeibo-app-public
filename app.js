let entries = JSON.parse(localStorage.getItem('kakeibo4') || '[]');
let viewMonth = new Date();
viewMonth.setDate(1);

let gMonthly = null;
let gPay = null;
let gBS = null;
let gPL = null;
let gCostType = null;
let editingId = null;
let editingBudgetEventId = null;
let uiPrefs = JSON.parse(localStorage.getItem('kakeibo_ui_prefs') || '{}');
let syncSettings = loadSyncSettings();
let monthlyClosings = loadMonthlyClosings();
let monthlyBudgets = loadMonthlyBudgets();
let budgetEvents = loadBudgetEvents();
let sessionSyncPassphrase = '';
let budgetMonth = new Date();
budgetMonth.setDate(1);

const defaultAccounts = {
  asset: ['現金','交通・電子マネー','SBI新生銀行','住信SBIネット銀行','ゆうちょ銀行','三井住友銀行','楽天銀行','中国銀行','NISA','固定資産','その他資産','ポイント'],
  liability: ['クレジットカード','奨学金','Paidy','消費者金融','その他負債'],
  income: ['給与','賞与','配当金','雑収入','銀行利息','プラス帳尻合わせ'],
  expense: ['食費','日用品費','家賃','水道代','ガス代','電気代','交通費','通信費','娯楽費','外食費','自己投資','沙奈費','交際費','旅費','被服費','美容費','保険','医療費','特別費','生活費','雑費','仕送り','税金等','マイナス帳尻合わせ']
};

const defaultExpenseColors = {
  食費:'#378ADD',
  日用品費:'#1D9E75',
  家賃:'#533AB7',
  水道代:'#185FA5',
  ガス代:'#BA7517',
  電気代:'#EF9F27',
  交通費:'#0F6E56',
  通信費:'#D4537E',
  娯楽費:'#D85A30',
  外食費:'#993C1D',
  自己投資:'#639922',
  沙奈費:'#3C3489',
  交際費:'#5DCAA5',
  旅費:'#7F77DD',
  被服費:'#ED93B1',
  美容費:'#D4537E',
  保険:'#888780',
  医療費:'#E24B4A',
  特別費:'#FA9F27',
  生活費:'#1D9E75',
  雑費:'#B4B2A9',
  仕送り:'#533AB7',
  税金等:'#444441',
  'マイナス帳尻合わせ':'#5F5E5A'
};

const defaultExpenseCostTypes = {
  食費:'variable',
  日用品費:'variable',
  家賃:'fixed',
  水道代:'fixed',
  ガス代:'fixed',
  電気代:'fixed',
  交通費:'variable',
  通信費:'fixed',
  娯楽費:'variable',
  外食費:'variable',
  自己投資:'variable',
  沙奈費:'variable',
  交際費:'variable',
  旅費:'variable',
  被服費:'variable',
  美容費:'variable',
  保険:'fixed',
  医療費:'variable',
  特別費:'variable',
  生活費:'variable',
  雑費:'variable',
  仕送り:'fixed',
  税金等:'fixed',
  'マイナス帳尻合わせ':'variable'
};

const POINT_ASSET_ACCOUNT = 'ポイント';
const POINT_ADJUSTMENT_INCOME = 'プラス帳尻合わせ';
const POINT_ROLE_ADJUSTMENT = 'point-adjustment';
const POINT_ROLE_EXPENSE = 'point-expense';
const OPENING_BALANCE_EQUITY = '開始残高調整';

let accountSettings = loadAccountSettings();

const payColors = {
  現金:'#888780',
  '交通・電子マネー':'#533AB7',
  SBI新生銀行:'#378ADD',
  住信SBIネット銀行:'#185FA5',
  ゆうちょ銀行:'#E24B4A',
  三井住友銀行:'#1D9E75',
  楽天銀行:'#D85A30',
  中国銀行:'#BA7517',
  ポイント:'#EF9F27',
  クレジットカード:'#3C3489',
  Paidy:'#D4537E'
};

let currentPreset = 'expense';

function randomColor() {
  const palette = ['#378ADD','#1D9E75','#533AB7','#185FA5','#BA7517','#EF9F27','#0F6E56','#D4537E','#D85A30','#639922','#7F77DD','#ED93B1','#888780','#E24B4A','#5F5E5A'];
  return palette[Math.floor(Math.random() * palette.length)];
}

function normalizeAccountBlock(items, type){
  if (!Array.isArray(items)) return [];
  return items
    .map(item => {
      if (typeof item === 'string') {
        return type === 'expense'
          ? { name:item, active:true, color:defaultExpenseColors[item] || randomColor(), costType:defaultExpenseCostTypes[item] || 'variable' }
          : { name:item, active:true };
      }
      if (item && typeof item.name === 'string') {
        const normalized = {
          name:item.name.trim(),
          active:item.active !== false
        };
        if (type === 'expense') {
          normalized.color = item.color || defaultExpenseColors[item.name] || randomColor();
          normalized.costType = item.costType === 'fixed' ? 'fixed' : (defaultExpenseCostTypes[item.name] || 'variable');
        }
        return normalized;
      }
      return null;
    })
    .filter(Boolean)
    .filter(item => item.name);
}

function loadAccountSettings(){
  const raw = JSON.parse(localStorage.getItem('kakeibo_accounts') || 'null');
  if (!raw) {
    return {
      asset: defaultAccounts.asset.map(name => ({ name, active:true })),
      liability: defaultAccounts.liability.map(name => ({ name, active:true })),
      income: defaultAccounts.income.map(name => ({ name, active:true })),
      expense: defaultAccounts.expense.map(name => ({
        name,
        active:true,
        color: defaultExpenseColors[name] || randomColor(),
        costType: defaultExpenseCostTypes[name] || 'variable'
      }))
    };
  }
  return {
    asset: ensureAccount(normalizeAccountBlock(raw.asset || defaultAccounts.asset, 'asset'), 'asset', POINT_ASSET_ACCOUNT),
    liability: normalizeAccountBlock(raw.liability || defaultAccounts.liability, 'liability'),
    income: ensureAccount(normalizeAccountBlock(raw.income || defaultAccounts.income, 'income'), 'income', POINT_ADJUSTMENT_INCOME),
    expense: normalizeAccountBlock(raw.expense || defaultAccounts.expense, 'expense')
  };
}

function ensureAccount(items, type, name) {
  const normalized = Array.isArray(items) ? [...items] : [];
  const existing = normalized.find(item => item.name === name);
  if (existing) {
    existing.active = true;
  } else {
    const account = { name, active:true };
    if (type === 'expense') {
      account.color = defaultExpenseColors[name] || randomColor();
      account.costType = defaultExpenseCostTypes[name] || 'variable';
    }
    normalized.push(account);
  }
  return normalized;
}

function saveAccountSettings(){
  localStorage.setItem('kakeibo_accounts', JSON.stringify(accountSettings));
  markLocalChanged();
}

function getAccounts(type, includeInactive = false){
  const list = accountSettings[type] || [];
  return includeInactive ? list.map(a => a.name) : list.filter(a => a.active).map(a => a.name);
}

function getExpenseColor(name){
  const found = (accountSettings.expense || []).find(a => a.name === name);
  return found?.color || defaultExpenseColors[name] || '#888';
}

function getExpenseCostType(name) {
  const found = (accountSettings.expense || []).find(a => a.name === name);
  return found?.costType === 'fixed' ? 'fixed' : 'variable';
}

function hasAccount(type, name){
  return (accountSettings[type] || []).some(a => a.name === name);
}

function isAccountUsed(type, name){
  if (type === 'asset' || type === 'expense') {
    if (entries.some(e => e.drCat === name)) return true;
  }
  if (type === 'asset' || type === 'liability' || type === 'income') {
    if (entries.some(e => e.crCat === name)) return true;
  }
  return false;
}

function getPresetConfig(){
  return {
    expense: {
      drLabel:'借方（費目）',
      crLabel:'貸方（支払方法）',
      hint:'食費・日用品費などをクレカや現金で支払う場合。',
      drOpts:[
        { g:'費用', opts:getAccounts('expense') },
        { g:'資産', opts:getAccounts('asset') }
      ],
      crOpts:[
        { g:'資産（支払元）', opts:getAccounts('asset') },
        { g:'負債', opts:getAccounts('liability') }
      ],
      drDef:getAccounts('expense')[0] || '',
      crDef:getAccounts('liability')[0] || getAccounts('asset')[0] || ''
    },
    income: {
      drLabel:'借方（受取口座）',
      crLabel:'貸方（収入科目）',
      hint:'給与・賞与などが口座に振り込まれる場合。',
      drOpts:[{ g:'資産', opts:getAccounts('asset') }],
      crOpts:[{ g:'収入', opts:getAccounts('income') }],
      drDef:getAccounts('asset')[0] || '',
      crDef:getAccounts('income')[0] || ''
    },
    repay: {
      drLabel:'借方（返済する負債）',
      crLabel:'貸方（支払い元の口座）',
      hint:'先月のクレカ請求などを口座から支払う場合。負債が減り、資産も減ります。',
      drOpts:[{ g:'負債', opts:getAccounts('liability') }],
      crOpts:[{ g:'資産（支払元）', opts:getAccounts('asset') }],
      drDef:getAccounts('liability')[0] || '',
      crDef:getAccounts('asset')[0] || ''
    },
    transfer: {
      drLabel:'借方（振替先）',
      crLabel:'貸方（振替元）',
      hint:'口座間の振替・現金引き出し・電子マネーチャージなど。',
      drOpts:[{ g:'資産', opts:getAccounts('asset') }],
      crOpts:[{ g:'資産', opts:getAccounts('asset') }],
      drDef:getAccounts('asset')[0] || '',
      crDef:getAccounts('asset')[1] || getAccounts('asset')[0] || ''
    },
    point: {
      drLabel:'借方（費目）',
      crLabel:'貸方（支払方法）',
      hint:'ポイント払いは「ポイント発生」と「ポイントから費用への支払い」をペアで自動記録します。',
      drOpts:[{ g:'費用', opts:getAccounts('expense') }],
      crOpts:[{ g:'資産', opts:getAccounts('asset') }],
      drDef:getAccounts('expense')[0] || '',
      crDef:getAccounts('asset').includes('ポイント') ? 'ポイント' : (getAccounts('asset')[0] || '')
    },
    opening: {
      drLabel:'借方（増やす側）',
      crLabel:'貸方（相手側）',
      hint:`使い始め時の初期残高を登録します。資産の開始残高は「資産 / ${OPENING_BALANCE_EQUITY}」、負債の開始残高は「${OPENING_BALANCE_EQUITY} / 負債」で入力します。月次収支には含めません。`,
      drOpts:[
        { g:'資産', opts:getAccounts('asset') },
        { g:'純資産（開始残高）', opts:[OPENING_BALANCE_EQUITY] }
      ],
      crOpts:[
        { g:'純資産（開始残高）', opts:[OPENING_BALANCE_EQUITY] },
        { g:'負債', opts:getAccounts('liability') }
      ],
      drDef:getAccounts('asset')[0] || OPENING_BALANCE_EQUITY,
      crDef:OPENING_BALANCE_EQUITY
    }
  };
}

function buildSelect(selId, optsGroups, defaultVal) {
  const sel = document.getElementById(selId);
  sel.innerHTML = optsGroups
    .map(g => {
      const opts = (g.opts || []).filter(Boolean);
      if (!opts.length) return '';
      return `<optgroup label="${g.g}">
        ${opts.map(o => `<option${o === defaultVal ? ' selected' : ''}>${o}</option>`).join('')}
      </optgroup>`;
    })
    .join('');

  if (!sel.innerHTML) {
    sel.innerHTML = '<option value="">科目がありません</option>';
  }
}

function setPreset(key) {
  currentPreset = key;
  document.querySelectorAll('.preset-btn').forEach(b => b.classList.remove('sel'));
  const activeBtn = document.getElementById('pr-' + key);
  if (activeBtn) activeBtn.classList.add('sel');

  const PRESETS = getPresetConfig();
  const p = PRESETS[key];
  document.getElementById('f-hint').textContent = p.hint;
  document.getElementById('f-dr-label').textContent = p.drLabel;
  document.getElementById('f-cr-label').textContent = p.crLabel;
  buildSelect('f-dr', p.drOpts, p.drDef);
  buildSelect('f-cr', p.crOpts, p.crDef);

  const eventSelect = document.getElementById('f-event');
  const eventHint = document.getElementById('f-event-hint');
  const isOpeningPreset = key === 'opening';
  if (eventSelect) {
    eventSelect.disabled = isOpeningPreset;
    if (isOpeningPreset) eventSelect.value = '';
  }
  if (eventHint && isOpeningPreset) {
    eventHint.textContent = '開始残高ではイベント予算は使いません。';
    eventHint.style.color = 'var(--text2)';
  } else {
    renderEventHint();
  }

  renderQuickCreditButtons();
}

function saveEntries() {
  localStorage.setItem('kakeibo4', JSON.stringify(entries));
  markLocalChanged();
}

function loadMonthlyClosings() {
  const raw = JSON.parse(localStorage.getItem('kakeibo_monthly_closings') || '[]');
  if (!Array.isArray(raw)) return [];
  return raw.map(normalizeMonthlyClosing).filter(Boolean);
}

function saveMonthlyClosings() {
  localStorage.setItem('kakeibo_monthly_closings', JSON.stringify(monthlyClosings));
  markLocalChanged();
}

function loadMonthlyBudgets() {
  const raw = JSON.parse(localStorage.getItem('kakeibo_monthly_budgets') || '{}');
  if (!raw || typeof raw !== 'object' || Array.isArray(raw)) return {};
  return Object.fromEntries(
    Object.entries(raw)
      .filter(([month]) => /^\d{4}-\d{2}$/.test(month))
      .map(([month, value]) => [month, normalizeMonthlyBudget(value)])
  );
}

function saveMonthlyBudgets() {
  localStorage.setItem('kakeibo_monthly_budgets', JSON.stringify(monthlyBudgets));
  markLocalChanged();
}

function loadBudgetEvents() {
  const raw = JSON.parse(localStorage.getItem('kakeibo_budget_events') || '[]');
  if (!Array.isArray(raw)) return [];
  return raw.map(normalizeBudgetEvent).filter(Boolean);
}

function saveBudgetEvents() {
  localStorage.setItem('kakeibo_budget_events', JSON.stringify(budgetEvents));
  markLocalChanged();
}

function saveUiPrefs() {
  localStorage.setItem('kakeibo_ui_prefs', JSON.stringify(uiPrefs));
}

function setQuickDate(offsetDays) {
  const d = new Date();
  d.setDate(d.getDate() + offsetDays);
  document.getElementById('f-date').value = formatLocalDate(d);
}

function applyLastUsedForm() {
  if (!uiPrefs.lastPreset && !uiPrefs.lastCreditByPreset) {
    alert('前回入力の記録がありません');
    return;
  }

  const preset = uiPrefs.lastPreset || 'expense';
  setPreset(preset);

  const rememberedCredit = uiPrefs.lastCreditByPreset?.[preset];
  if (rememberedCredit) {
    const cr = document.getElementById('f-cr');
    if ([...cr.options].some(o => o.value === rememberedCredit)) {
      cr.value = rememberedCredit;
    }
  }

  const rememberedDebit = uiPrefs.lastDebitByPreset?.[preset];
  if (rememberedDebit) {
    const dr = document.getElementById('f-dr');
    if ([...dr.options].some(o => o.value === rememberedDebit)) {
      dr.value = rememberedDebit;
    }
  }

  renderQuickCreditButtons();
}

function updateListCategoryFilterOptions() {
  const sel = document.getElementById('list-category-filter');
  if (!sel) return;

  const current = sel.value;
  const categories = Array.from(new Set([
    ...getAccounts('expense', true),
    ...getAccounts('income', true),
    ...getAccounts('asset', true),
    ...getAccounts('liability', true)
  ])).sort((a, b) => a.localeCompare(b, 'ja'));

  sel.innerHTML = `<option value="">すべての科目</option>` +
    categories.map(c => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join('');

  if (categories.includes(current)) {
    sel.value = current;
  }
}

function updateListEventFilterOptions() {
  const sel = document.getElementById('list-event-filter');
  if (!sel) return;

  const current = sel.value;
  sel.innerHTML = '<option value="">すべてのイベント</option>' +
    budgetEvents
      .slice()
      .sort((a, b) => String(a.startDate || '').localeCompare(String(b.startDate || '')) || a.name.localeCompare(b.name, 'ja'))
      .map(event => `<option value="${escapeHtml(event.id)}">${escapeHtml(event.name)}</option>`)
      .join('');

  if (budgetEvents.some(event => event.id === current)) {
    sel.value = current;
  }
}

function renderEventOptions(selectedId = '') {
  const sel = document.getElementById('f-event');
  if (!sel) return;
  const dateValue = document.getElementById('f-date')?.value || '';
  const matches = [];
  const others = [];

  budgetEvents
    .slice()
    .sort((a, b) => String(a.startDate || '').localeCompare(String(b.startDate || '')) || a.name.localeCompare(b.name, 'ja'))
    .forEach(event => {
      if (isEventDateMatch(event, dateValue)) matches.push(event);
      else others.push(event);
    });

  const eventOption = event => {
    const period = event.startDate || event.endDate
      ? ` (${event.startDate || '未設定'} - ${event.endDate || '未設定'})`
      : '';
    return `<option value="${escapeHtml(event.id)}">${escapeHtml(event.name + period)}</option>`;
  };

  const parts = ['<option value="">イベントなし</option>'];
  if (matches.length) {
    parts.push(`<optgroup label="日付に合うイベント">${matches.map(eventOption).join('')}</optgroup>`);
  }
  if (others.length) {
    parts.push(`<optgroup label="その他のイベント">${others.map(eventOption).join('')}</optgroup>`);
  }
  sel.innerHTML = parts.join('');

  const availableIds = budgetEvents.map(event => event.id);
  if (selectedId && availableIds.includes(selectedId)) {
    sel.value = selectedId;
  } else {
    const suggestedId = getSuggestedEventId(dateValue);
    sel.value = suggestedId || '';
  }

  renderEventHint();
}

function isEventDateMatch(event, dateValue) {
  if (!dateValue) return false;
  if (event.startDate && dateValue < event.startDate) return false;
  if (event.endDate && dateValue > event.endDate) return false;
  return true;
}

function getSuggestedEventId(dateValue) {
  if (!dateValue) return '';
  const matched = budgetEvents.filter(event => isEventDateMatch(event, dateValue));
  return matched.length === 1 ? matched[0].id : '';
}

function renderEventHint() {
  const hint = document.getElementById('f-event-hint');
  const sel = document.getElementById('f-event');
  const dateValue = document.getElementById('f-date')?.value || '';
  if (!hint || !sel) return;
  if (currentPreset === 'opening') {
    hint.innerHTML = '開始残高ではイベント予算は使いません。';
    hint.style.color = 'var(--text2)';
    return;
  }

  const selected = getBudgetEvent(sel.value);
  const matched = budgetEvents.filter(event => isEventDateMatch(event, dateValue));

  if (selected) {
    if (dateValue && !isEventDateMatch(selected, dateValue)) {
      hint.innerHTML = `選択中のイベント期間は ${escapeHtml(selected.startDate || '未設定')} - ${escapeHtml(selected.endDate || '未設定')} です。入力日が期間外です。`;
      hint.style.color = 'var(--red-dark)';
      return;
    }
    const spent = getEventSpend(selected.id);
    hint.innerHTML = `選択中: ${escapeHtml(selected.name)} / 予算 ${fmt(selected.budget)} / 使用額 ${fmt(spent)}`;
    hint.style.color = 'var(--text2)';
    return;
  }

  if (matched.length === 1) {
    hint.innerHTML = `この日付には ${escapeHtml(matched[0].name)} が候補です。`;
    hint.style.color = 'var(--text2)';
    return;
  }
  if (matched.length > 1) {
    hint.innerHTML = `この日付に合うイベントが ${matched.length} 件あります。`;
    hint.style.color = 'var(--text2)';
    return;
  }

  hint.innerHTML = 'イベントを使わない支出は、そのまま「イベントなし」で記録できます。';
  hint.style.color = 'var(--text2)';
}

function getQuickCreditCandidates() {
  const preset = currentPreset;
  const remembered = uiPrefs.lastCreditByPreset?.[preset];
  const allAsset = getAccounts('asset');
  const allLiability = getAccounts('liability');

  let candidates = [];

  if (preset === 'expense') {
    candidates = [...allLiability, ...allAsset];
  } else if (preset === 'income') {
    candidates = getAccounts('income');
  } else if (preset === 'repay') {
    candidates = allAsset;
  } else if (preset === 'transfer') {
    candidates = allAsset;
  } else if (preset === 'point') {
    candidates = allAsset;
  } else if (preset === 'opening') {
    candidates = [OPENING_BALANCE_EQUITY, ...allLiability];
  }

  const uniq = Array.from(new Set(candidates.filter(Boolean)));
  if (remembered && uniq.includes(remembered)) {
    return [remembered, ...uniq.filter(v => v !== remembered)].slice(0, 6);
  }
  return uniq.slice(0, 6);
}

function renderQuickCreditButtons() {
  const wrap = document.getElementById('quick-credit-wrap');
  if (!wrap) return;

  const cr = document.getElementById('f-cr');
  const current = cr?.value || '';
  const candidates = getQuickCreditCandidates();

  if (!candidates.length) {
    wrap.innerHTML = '';
    return;
  }

  wrap.innerHTML = candidates.map(name => `
    <button
      type="button"
      class="quick-credit-btn ${name === current ? 'active' : ''}"
      onclick="selectQuickCredit('${escapeJs(name)}')"
    >${escapeHtml(name)}</button>
  `).join('');
}

function selectQuickCredit(name) {
  const cr = document.getElementById('f-cr');
  if (!cr) return;
  if ([...cr.options].some(o => o.value === name)) {
    cr.value = name;
    renderQuickCreditButtons();
  }
}

function moveFocusAfterAdd() {
  document.getElementById('f-amt').focus();
}

function formatLocalDate(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

function setupFormInteractions() {
  const amt = document.getElementById('f-amt');
  const desc = document.getElementById('f-desc');
  const dr = document.getElementById('f-dr');
  const cr = document.getElementById('f-cr');
  const date = document.getElementById('f-date');
  const event = document.getElementById('f-event');

  if (amt && !amt.dataset.bound) {
    amt.addEventListener('keydown', e => {
      if (e.key === 'Enter') {
        e.preventDefault();
        desc.focus();
      }
    });
    amt.dataset.bound = '1';
  }

  if (desc && !desc.dataset.bound) {
    desc.addEventListener('keydown', e => {
      if (e.key === 'Enter') {
        e.preventDefault();
        addEntry();
      }
    });
    desc.dataset.bound = '1';
  }

  if (dr && !dr.dataset.bound) {
    dr.addEventListener('change', () => {
      if (!document.getElementById('f-desc').value.trim()) {
        const selectedText = dr.value || '';
        document.getElementById('f-desc').placeholder = selectedText ? `例：${selectedText}` : '例：スーパーで買い物';
      }
    });
    dr.dataset.bound = '1';
  }

  if (cr && !cr.dataset.bound) {
    cr.addEventListener('change', () => {
      renderQuickCreditButtons();
    });
    cr.dataset.bound = '1';
  }

  if (date && !date.dataset.bound) {
    date.addEventListener('change', () => {
      const currentEventId = document.getElementById('f-event')?.value || '';
      const keepCurrent = currentEventId && isEventDateMatch(getBudgetEvent(currentEventId) || {}, date.value);
      renderEventOptions(keepCurrent ? currentEventId : '');
    });
    date.dataset.bound = '1';
  }

  if (event && !event.dataset.bound) {
    event.addEventListener('change', () => {
      renderEventHint();
    });
    event.dataset.bound = '1';
  }
}

function fmt(n) {
  return '¥' + Math.round(Math.abs(n)).toLocaleString();
}

function fmtSigned(n) {
  if (n === 0) return '¥0';
  return `${n < 0 ? '-' : '+'}¥${Math.round(Math.abs(n)).toLocaleString()}`;
}

function fmtS(n) {
  if (n === 0) return '<span class="zero">—</span>';
  const s = Math.round(Math.abs(n)).toLocaleString();
  return n < 0 ? `<span class="neg">-¥${s}</span>` : `¥${s}`;
}

function fmtSC(n) {
  if (n === 0) return '<span class="zero">—</span>';
  const s = Math.round(Math.abs(n)).toLocaleString();
  const c = n >= 0 ? 'pos' : 'neg';
  return `<span class="${c}">${n < 0 ? '-' : ''}¥${s}</span>`;
}

function getMonths(n) {
  const now = new Date();
  const res = [];
  for (let i = n - 1; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    res.push(d);
  }
  return res;
}

function mLabel(d) {
  return `${d.getMonth() + 1}月`;
}

function mEntries(d) {
  const y = d.getFullYear();
  const m = d.getMonth();
  return entries.filter(e => {
    const ed = new Date(e.date);
    return ed.getFullYear() === y && ed.getMonth() === m;
  });
}

function nowEntries() {
  return mEntries(new Date());
}

function isOpeningEntry(entry) {
  return entry?.entryType === 'opening' || entry?.preset === 'opening';
}

function isIncome(e) {
  return !isOpeningEntry(e) && getAccounts('income', true).includes(e.crCat);
}

function isExpense(e) {
  return !isOpeningEntry(e) && getAccounts('expense', true).includes(e.drCat);
}

function isOpeningAssetEntry(entry) {
  return isOpeningEntry(entry) && getAccounts('asset', true).includes(entry.drCat) && entry.crCat === OPENING_BALANCE_EQUITY;
}

function isOpeningLiabilityEntry(entry) {
  return isOpeningEntry(entry) && entry.drCat === OPENING_BALANCE_EQUITY && getAccounts('liability', true).includes(entry.crCat);
}

function isValidOpeningEntry(entry) {
  return isOpeningAssetEntry(entry) || isOpeningLiabilityEntry(entry);
}

function acctFlow(name, month){
  let t = 0;
  mEntries(month).forEach(e => {
    if (e.drCat === name) t += e.amount;
    if (e.crCat === name) t -= e.amount;
  });
  return t;
}

function acctCumul(name, upToMonth){
  let t = 0;
  const upTo = new Date(upToMonth.getFullYear(), upToMonth.getMonth()+1, 0, 23, 59, 59);
  entries.forEach(e => {
    const d = new Date(e.date);
    if (d > upTo) return;
    if (e.drCat === name) t += e.amount;
    if (e.crCat === name) t -= e.amount;
  });
  return t;
}

function monthKeyFromDate(dateValue) {
  const d = new Date(dateValue);
  if (Number.isNaN(d.getTime())) return '';
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
}

function monthKeyFromMonth(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
}

function formatMonthLabel(date) {
  return `${date.getFullYear()}年${date.getMonth() + 1}月`;
}

function getMonthlyBudget(monthKey) {
  return monthlyBudgets[monthKey] || {};
}

function getBudgetEvent(eventId) {
  return budgetEvents.find(event => event.id === eventId) || null;
}

function getEventSpend(eventId) {
  return entries
    .filter(e => e.eventId === eventId && isExpense(e))
    .reduce((sum, e) => sum + e.amount, 0);
}

function getEventBreakdown(eventId, field) {
  const map = {};
  entries
    .filter(e => e.eventId === eventId && isExpense(e))
    .forEach(entry => {
      const key = entry[field] || '未設定';
      map[key] = (map[key] || 0) + entry.amount;
    });

  return Object.entries(map)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);
}

function getMonthBudgetTotals(monthDate) {
  const monthKey = monthKeyFromMonth(monthDate);
  const budget = getMonthlyBudget(monthKey);
  const monthEntries = mEntries(monthDate).filter(isExpense);
  const actualByCategory = {};
  monthEntries.forEach(entry => {
    actualByCategory[entry.drCat] = (actualByCategory[entry.drCat] || 0) + entry.amount;
  });

  const totals = {
    budget:0,
    actual:0,
    fixedBudget:0,
    fixedActual:0,
    variableBudget:0,
    variableActual:0
  };

  getAccounts('expense', true).forEach(name => {
    const budgetValue = Number(budget[name] || 0);
    const actualValue = Number(actualByCategory[name] || 0);
    const costType = getExpenseCostType(name);

    totals.budget += budgetValue;
    totals.actual += actualValue;
    if (costType === 'fixed') {
      totals.fixedBudget += budgetValue;
      totals.fixedActual += actualValue;
    } else {
      totals.variableBudget += budgetValue;
      totals.variableActual += actualValue;
    }
  });

  return totals;
}

function getClosing(monthKey) {
  return monthlyClosings.find(c => c.month === monthKey) || null;
}

function isMonthClosed(monthKey) {
  return Boolean(getClosing(monthKey));
}

function blockIfClosedMonth(monthKey, actionLabel) {
  if (!monthKey || !isMonthClosed(monthKey)) return false;
  alert(`${monthKey} は締め済みです。${actionLabel}するには、財務タブで締めを解除してください。`);
  return true;
}

function getActiveTab() {
  const tabs = ['record','list','graph','budget','summary','fs','settings'];
  return tabs.find(t => document.getElementById('tb-' + t)?.classList.contains('active')) || 'record';
}

function refreshActiveTab() {
  updateMetrics();
  const active = getActiveTab();
  if (active === 'list') renderList();
  if (active === 'graph') renderGraph();
  if (active === 'budget') renderBudget();
  if (active === 'summary') renderSummary();
  if (active === 'fs') {
    if (!document.getElementById('fs-month').options.length) buildFSMonthOptions();
    renderFS();
  }
  if (active === 'settings') renderSettings();
}

function refreshAccountDrivenUI(){
  const editing = editingId ? entries.find(e => e.id === editingId) : null;
  if (editing) {
    setPreset(editing.preset || guessPreset(editing));
    document.getElementById('f-dr').value = editing.drCat || '';
    document.getElementById('f-cr').value = editing.crCat || '';
    renderEventOptions(editing.eventId || '');
  } else {
    setPreset(uiPrefs.lastPreset || currentPreset);
  }
  renderSettings();
  updateMetrics();
  updateListCategoryFilterOptions();
  updateListEventFilterOptions();
  renderQuickCreditButtons();
  renderBudget();
}

function updateMetrics() {
  const es = nowEntries();
  const inc = es.filter(isIncome).reduce((s, e) => s + e.amount, 0);
  const exp = es.filter(isExpense).reduce((s, e) => s + e.amount, 0);

  document.getElementById('m-inc').textContent = fmt(inc);
  document.getElementById('m-exp').textContent = fmt(exp);

  const bal = inc - exp;
  const b = document.getElementById('m-bal');
  b.textContent = fmtSigned(bal);
  b.style.color = bal >= 0 ? 'var(--green)' : 'var(--red)';
}

function renderList() {
  updateListCategoryFilterOptions();
  updateListEventFilterOptions();

  const keyword = (document.getElementById('list-search')?.value || '').trim().toLowerCase();
  const categoryFilter = document.getElementById('list-category-filter')?.value || '';
  const eventFilter = document.getElementById('list-event-filter')?.value || '';
  const filteredEvent = eventFilter ? getBudgetEvent(eventFilter) : null;
  const monthNav = document.querySelector('#t-list .mnav');

  if (filteredEvent) {
    document.getElementById('list-lbl').textContent = `${filteredEvent.name} の取引`;
    if (monthNav) monthNav.style.display = 'none';
  } else {
    const y = viewMonth.getFullYear();
    const m = viewMonth.getMonth();
    document.getElementById('list-lbl').textContent = `${y}年${m + 1}月`;
    if (monthNav) monthNav.style.display = '';
  }

  const sourceEntries = filteredEvent
    ? entries.filter(e => e.eventId === filteredEvent.id)
    : mEntries(viewMonth);

  const es = sourceEntries
    .filter(e => {
      const matchesKeyword = !keyword || [
        e.desc || '',
        e.drCat || '',
        e.crCat || '',
        e.drNote || '',
        e.crNote || ''
      ].join(' ').toLowerCase().includes(keyword);

      const matchesCategory = !categoryFilter || e.drCat === categoryFilter || e.crCat === categoryFilter;

      return matchesKeyword && matchesCategory;
    })
    .sort((a, b) => {
      const dateDiff = new Date(b.date) - new Date(a.date);
      if (dateDiff !== 0) return dateDiff;
      return String(b.id).localeCompare(String(a.id));
    });

  const el = document.getElementById('elist');
  if (!es.length) {
    el.innerHTML = '<div class="empty">条件に合う記録はありません</div>';
    return;
  }

  el.innerHTML = es.map(e => `
    <div class="entry ${isPointAdjustmentEntry(e) ? 'entry-auto' : ''}">
      <div class="etop">
        <div>
          <div class="ememo">${escapeHtml(e.desc || e.drCat)}</div>
          <div class="edate">${escapeHtml(e.date)}${e.eventId ? ` · ${escapeHtml(getBudgetEvent(e.eventId)?.name || 'イベント')}` : ''}${isPointLinkedEntry(e) ? ' · ポイント払い' : ''}${isPointAdjustmentEntry(e) ? ' · 自動補助' : ''}${isOpeningEntry(e) ? ' · 開始残高' : ''}</div>
        </div>
        <div class="eright-actions">
          <div class="eamt">${fmt(e.amount)}</div>
          <div class="eright">
            <button class="editbtn" onclick="startEdit('${e.id}')">編集</button>
            <button class="edel" onclick="delEntry('${e.id}')">×</button>
          </div>
        </div>
      </div>
      <div class="elegs">
        <div class="leg dr">
          <div class="leglbl">借方</div>
          ${escapeHtml(e.drCat)}${e.drNote ? ' · ' + escapeHtml(e.drNote) : ''}
        </div>
        <div class="arrmid">→</div>
        <div class="leg cr">
          <div class="leglbl">貸方</div>
          ${escapeHtml(e.crCat)}${e.crNote ? ' · ' + escapeHtml(e.crNote) : ''}
        </div>
      </div>
    </div>
  `).join('');
}

function openEventEntries(eventId) {
  const event = getBudgetEvent(eventId);
  if (!event) return;

  const eventFilter = document.getElementById('list-event-filter');
  const search = document.getElementById('list-search');
  const categoryFilter = document.getElementById('list-category-filter');

  if (event.startDate) {
    const d = new Date(event.startDate);
    if (!Number.isNaN(d.getTime())) {
      viewMonth = new Date(d.getFullYear(), d.getMonth(), 1);
    }
  }

  if (search) search.value = '';
  if (categoryFilter) categoryFilter.value = '';
  updateListEventFilterOptions();
  if (eventFilter) eventFilter.value = eventId;
  sw('list');
}

function delEntry(id) {
  const entry = entries.find(e => e.id === id);
  const targets = getLinkedPointEntries(entry);
  if (targets.some(e => blockIfClosedMonth(monthKeyFromDate(e.date), '削除'))) return;
  const message = targets.length > 1
    ? 'ポイント払いのペア仕訳をまとめて削除しますか？'
    : '削除しますか？';
  if (!confirm(message)) return;
  const targetIds = new Set(targets.map(e => e.id));
  entries = entries.filter(e => !targetIds.has(e.id));
  saveEntries();
  if (targetIds.has(editingId)) cancelEdit(false);
  refreshActiveTab();
  renderSettings();
}

function chMon(d) {
  viewMonth.setMonth(viewMonth.getMonth() + d);
  renderList();
}

function renderGraph() {
  const months = getMonths(6);
  const labels = months.map(mLabel);

  const incD = months.map(m => mEntries(m).filter(isIncome).reduce((s, e) => s + e.amount, 0));
  const expD = months.map(m => mEntries(m).filter(isExpense).reduce((s, e) => s + e.amount, 0));

  if (gMonthly) gMonthly.destroy();
  gMonthly = new Chart(document.getElementById('gc-monthly'), {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label:'収入', data:incD, backgroundColor:'rgba(29,158,117,0.7)', borderRadius:3 },
        { label:'支出', data:expD, backgroundColor:'rgba(216,90,48,0.7)', borderRadius:3 }
      ]
    },
    options: {
      responsive:true,
      maintainAspectRatio:false,
      plugins:{ legend:{ display:false } },
      scales:{
        x:{ grid:{ display:false }, ticks:{ color:'#888', font:{ size:10 } } },
        y:{ grid:{ color:'rgba(128,128,128,0.12)' }, ticks:{ color:'#888', font:{ size:9 }, callback:v=>'¥'+v.toLocaleString() } }
      }
    }
  });

  const curExp = nowEntries().filter(isExpense);
  const cm = {};
  curExp.forEach(e => { cm[e.drCat] = (cm[e.drCat] || 0) + e.amount; });
  const ce = Object.entries(cm).sort((a,b) => b[1] - a[1]);
  const ct = ce.reduce((s,[,v]) => s + v, 0);

  document.getElementById('gc-catbars').innerHTML = ce.length
    ? ce.map(([c,v]) => `
      <div class="cbar-row">
        <div class="cbar-lbl" title="${escapeHtml(c)}">${escapeHtml(c)}</div>
        <div class="cbar-trk">
          <div class="cbar-fill" style="width:${ct ? Math.round(v/ct*100) : 0}%;background:${getExpenseColor(c)};"></div>
        </div>
        <div class="cbar-val">${fmt(v)}</div>
      </div>
    `).join('')
    : '<div class="empty" style="padding:1rem;">データなし</div>';

  const assetNames = getAccounts('asset', true);
  const liabilityNames = getAccounts('liability', true);
  const pm = {};
  curExp.forEach(e => {
    if ([...assetNames, ...liabilityNames].includes(e.crCat)) {
      pm[e.crCat] = (pm[e.crCat] || 0) + e.amount;
    }
  });
  const pe = Object.entries(pm).sort((a,b) => b[1] - a[1]);
  const pt = pe.reduce((s,[,v]) => s + v, 0);

  document.getElementById('gc-paylgd').innerHTML = pe.map(([p,v]) => `
    <span><span class="ldot" style="background:${payColors[p] || '#888'};"></span>${escapeHtml(p)} ${pt ? Math.round(v/pt*100) : 0}%</span>
  `).join('');

  if (gPay) { gPay.destroy(); gPay = null; }
  if (pe.length) {
    gPay = new Chart(document.getElementById('gc-pay'), {
      type:'doughnut',
      data:{
        labels:pe.map(([p]) => p),
        datasets:[{
          data:pe.map(([,v]) => v),
          backgroundColor:pe.map(([p]) => payColors[p] || '#888'),
          borderWidth:0
        }]
      },
      options:{ responsive:true, maintainAspectRatio:false, plugins:{ legend:{ display:false } }, cutout:'60%' }
    });
  }

  const fixedActual = curExp
    .filter(e => getExpenseCostType(e.drCat) === 'fixed')
    .reduce((sum, e) => sum + e.amount, 0);
  const variableActual = curExp
    .filter(e => getExpenseCostType(e.drCat) !== 'fixed')
    .reduce((sum, e) => sum + e.amount, 0);

  if (gCostType) gCostType.destroy();
  gCostType = new Chart(document.getElementById('gc-costtype'), {
    type:'bar',
    data:{
      labels:['固定費', '変動費'],
      datasets:[{
        data:[fixedActual, variableActual],
        backgroundColor:['rgba(83,58,183,0.75)', 'rgba(55,138,221,0.75)'],
        borderRadius:4
      }]
    },
    options:{
      responsive:true,
      maintainAspectRatio:false,
      plugins:{ legend:{ display:false } },
      scales:{
        x:{ grid:{ display:false }, ticks:{ color:'#888', font:{ size:10 } } },
        y:{ grid:{ color:'rgba(128,128,128,0.12)' }, ticks:{ color:'#888', font:{ size:9 }, callback:v => '¥' + v.toLocaleString() } }
      }
    }
  });
}

function renderBudget() {
  const monthKey = monthKeyFromMonth(budgetMonth);
  const budget = getMonthlyBudget(monthKey);
  const totals = getMonthBudgetTotals(budgetMonth);
  const actualByCategory = {};
  mEntries(budgetMonth)
    .filter(isExpense)
    .forEach(entry => {
      actualByCategory[entry.drCat] = (actualByCategory[entry.drCat] || 0) + entry.amount;
    });

  const monthLabel = document.getElementById('budget-month-label');
  if (monthLabel) monthLabel.textContent = `${formatMonthLabel(budgetMonth)} の予算`;

  const summary = document.getElementById('budget-month-summary');
  if (summary) {
    const totalOver = totals.budget > 0 && totals.actual > totals.budget;
    const fixedOver = totals.fixedBudget > 0 && totals.fixedActual > totals.fixedBudget;
    const variableOver = totals.variableBudget > 0 && totals.variableActual > totals.variableBudget;
    summary.innerHTML = `
      <div class="budget-stat"><span>予算合計</span><strong>${fmt(totals.budget)}</strong></div>
      <div class="budget-stat ${totalOver ? 'alert' : ''}"><span>実績合計</span><strong class="${totalOver ? 'neg' : 'pos'}">${fmt(totals.actual)}</strong></div>
      <div class="budget-stat ${fixedOver ? 'alert' : ''}"><span>固定費 予算 / 実績</span><strong>${fmt(totals.fixedBudget)} / ${fmt(totals.fixedActual)}</strong></div>
      <div class="budget-stat ${variableOver ? 'alert' : ''}"><span>変動費 予算 / 実績</span><strong>${fmt(totals.variableBudget)} / ${fmt(totals.variableActual)}</strong></div>
    `;
  }

  const list = document.getElementById('monthly-budget-list');
  if (list) {
    list.innerHTML = getAccounts('expense', true).map(name => {
      const budgetValue = Number(budget[name] || 0);
      const actualValue = Number(actualByCategory[name] || 0);
      const diff = budgetValue - actualValue;
      const costType = getExpenseCostType(name);
      const diffClass = diff === 0 ? 'zero' : (diff > 0 ? 'under' : 'over');
      const over = budgetValue > 0 && actualValue > budgetValue;
      const progress = budgetValue > 0 ? Math.min(100, Math.round(actualValue / budgetValue * 100)) : 0;

      return `
        <div class="budget-row ${over ? 'over' : ''}">
          <div class="budget-row-main">
            <div class="budget-row-title">
              <span class="color-dot" style="background:${getExpenseColor(name)};"></span>
              <span>${escapeHtml(name)}</span>
              <span class="pill ${costType}">${costType === 'fixed' ? '固定費' : '変動費'}</span>
            </div>
            <div class="budget-row-meta">実績 ${fmt(actualValue)}</div>
            ${budgetValue > 0 ? `
              <div class="budget-progress">
                <div class="budget-progress-fill ${over ? 'warn' : ''}" style="width:${progress}%;"></div>
              </div>
            ` : ''}
          </div>
          <input
            type="number"
            inputmode="numeric"
            min="0"
            value="${budgetValue || ''}"
            placeholder="予算"
            onchange="setMonthlyBudgetValue('${escapeJs(monthKey)}', '${escapeJs(name)}', this.value)"
          >
          <div class="budget-diff ${diffClass}">
            ${budgetValue > 0 ? `差額 ${fmtSigned(diff)}` : '予算未設定'}
          </div>
        </div>
      `;
    }).join('');
  }

  const eventList = document.getElementById('budget-event-list');
  if (eventList) {
    if (!budgetEvents.length) {
      eventList.innerHTML = '<div class="budget-empty">イベント予算はまだありません</div>';
    } else {
      eventList.innerHTML = budgetEvents
        .slice()
        .sort((a, b) => String(a.startDate || '').localeCompare(String(b.startDate || '')) || a.name.localeCompare(b.name, 'ja'))
        .map(event => {
          const spent = getEventSpend(event.id);
          const remaining = event.budget - spent;
          const linkedCount = entries.filter(entry => entry.eventId === event.id).length;
          const periodText = `${event.startDate || '開始日未設定'} - ${event.endDate || '終了日未設定'}`;
          const categoryBreakdown = getEventBreakdown(event.id, 'drCat');
          const paymentBreakdown = getEventBreakdown(event.id, 'crCat');
          const renderBreakdownRows = rows => rows.length
            ? rows.map(([name, value]) => `
                <div class="budget-breakdown-row">
                  <span>${escapeHtml(name)}</span>
                  <strong>${fmt(value)}</strong>
                </div>
              `).join('')
            : '<div class="budget-breakdown-empty">まだ支出がありません</div>';
          return `
            <div class="budget-event-card">
              <div class="budget-event-head">
                <div>
                  <div class="budget-event-title">${escapeHtml(event.name)}</div>
                  <div class="budget-event-sub ${remaining < 0 ? 'warn' : ''}">${escapeHtml(periodText)} / 紐づき ${linkedCount}件</div>
                </div>
                <span class="acct-tag ${remaining >= 0 ? 'on' : 'off'}">${remaining >= 0 ? '予算内' : '超過'}</span>
              </div>
              <div class="budget-event-metrics">
                <div class="budget-event-metric"><span>予算</span><strong>${fmt(event.budget)}</strong></div>
                <div class="budget-event-metric"><span>使用額</span><strong>${fmt(spent)}</strong></div>
                <div class="budget-event-metric"><span>残額</span><strong class="${remaining >= 0 ? 'pos' : 'neg'}">${fmtSigned(remaining)}</strong></div>
              </div>
              <div class="budget-breakdowns">
                <div class="budget-breakdown">
                  <div class="budget-breakdown-title">費目別内訳</div>
                  ${renderBreakdownRows(categoryBreakdown)}
                </div>
                <div class="budget-breakdown">
                  <div class="budget-breakdown-title">支払方法別内訳</div>
                  ${renderBreakdownRows(paymentBreakdown)}
                </div>
              </div>
              <div class="sub-actions">
                <button class="btnsub" type="button" onclick="openEventEntries('${escapeJs(event.id)}')">取引を見る</button>
                <button class="btnsub" type="button" onclick="editBudgetEvent('${escapeJs(event.id)}')">編集</button>
                <button class="btnsub" type="button" onclick="deleteBudgetEvent('${escapeJs(event.id)}')">削除</button>
              </div>
            </div>
          `;
        }).join('');
    }
  }

  renderEventOptions(document.getElementById('f-event')?.value || '');
}

function changeBudgetMonth(offset) {
  budgetMonth.setMonth(budgetMonth.getMonth() + offset);
  renderBudget();
}

function setMonthlyBudgetValue(monthKey, category, rawValue) {
  const value = Number(rawValue || 0);
  monthlyBudgets[monthKey] = monthlyBudgets[monthKey] || {};
  if (!Number.isFinite(value) || value <= 0) {
    delete monthlyBudgets[monthKey][category];
  } else {
    monthlyBudgets[monthKey][category] = Math.round(value);
  }

  if (!Object.keys(monthlyBudgets[monthKey]).length) {
    delete monthlyBudgets[monthKey];
  }

  saveMonthlyBudgets();
  renderBudget();
}

function copyPreviousMonthBudget() {
  const currentKey = monthKeyFromMonth(budgetMonth);
  const prev = new Date(budgetMonth.getFullYear(), budgetMonth.getMonth() - 1, 1);
  const prevKey = monthKeyFromMonth(prev);
  const source = getMonthlyBudget(prevKey);

  if (!Object.keys(source).length) {
    alert('前月にコピーできる予算がありません');
    return;
  }

  monthlyBudgets[currentKey] = { ...source };
  saveMonthlyBudgets();
  renderBudget();
  alert('前月の予算をコピーしました');
}

function clearBudgetMonth() {
  const monthKey = monthKeyFromMonth(budgetMonth);
  if (!monthlyBudgets[monthKey]) {
    alert('この月の予算はまだ設定されていません');
    return;
  }
  if (!confirm(`${monthKey} の予算をクリアしますか？`)) return;
  delete monthlyBudgets[monthKey];
  saveMonthlyBudgets();
  renderBudget();
}

function saveBudgetEvent() {
  const isEditing = Boolean(editingBudgetEventId);
  const name = (document.getElementById('budget-event-name')?.value || '').trim();
  const amount = Number(document.getElementById('budget-event-amount')?.value || 0);
  const startDate = document.getElementById('budget-event-start')?.value || '';
  const endDate = document.getElementById('budget-event-end')?.value || '';

  if (!name) {
    alert('イベント名を入力してください');
    return;
  }
  if (!Number.isFinite(amount) || amount < 0) {
    alert('予算額を入力してください');
    return;
  }
  if (startDate && endDate && startDate > endDate) {
    alert('終了日は開始日以降にしてください');
    return;
  }

  const payload = {
    id: editingBudgetEventId || newEntryId(),
    name,
    budget: Math.round(amount),
    startDate,
    endDate
  };

  budgetEvents = budgetEvents.filter(event => event.id !== payload.id);
  budgetEvents.push(payload);
  saveBudgetEvents();
  cancelBudgetEventEdit(false);
  renderBudget();
  alert(isEditing ? 'イベント予算を更新しました' : 'イベント予算を追加しました');
}

function editBudgetEvent(id) {
  const event = getBudgetEvent(id);
  if (!event) return;
  editingBudgetEventId = id;
  document.getElementById('budget-event-name').value = event.name || '';
  document.getElementById('budget-event-amount').value = event.budget ?? '';
  document.getElementById('budget-event-start').value = event.startDate || '';
  document.getElementById('budget-event-end').value = event.endDate || '';
  document.getElementById('budget-event-submit').textContent = 'イベント予算を更新';
}

function cancelBudgetEventEdit(showAlert = false) {
  editingBudgetEventId = null;
  const fields = ['budget-event-name', 'budget-event-amount', 'budget-event-start', 'budget-event-end'];
  fields.forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });
  const submit = document.getElementById('budget-event-submit');
  if (submit) submit.textContent = 'イベント予算を追加';
  if (showAlert) alert('イベント予算の編集をキャンセルしました');
}

function deleteBudgetEvent(id) {
  const event = getBudgetEvent(id);
  if (!event) return;
  const linkedCount = entries.filter(entry => entry.eventId === id).length;
  const message = linkedCount
    ? `「${event.name}」を削除すると、${linkedCount}件の取引からイベント紐づけも外れます。削除しますか？`
    : `「${event.name}」を削除しますか？`;
  if (!confirm(message)) return;

  budgetEvents = budgetEvents.filter(item => item.id !== id);
  entries = entries.map(entry => entry.eventId === id ? { ...entry, eventId:undefined } : entry);
  saveBudgetEvents();
  saveEntries();
  if (editingBudgetEventId === id) cancelBudgetEventEdit(false);
  renderBudget();
  refreshActiveTab();
}

function renderSummary() {
  const n = parseInt(document.getElementById('s-months').value, 10);
  const months = getMonths(n);

  const SCHEMA = [
    { key:'asset', label:'資産', badge:'ba', accts:getAccounts('asset', true), sign:1 },
    { key:'liab', label:'負債', badge:'bl', accts:getAccounts('liability', true), sign:-1 },
    { key:'equity', label:'純資産', badge:'be', computed:true, formula:(a,l)=>a-l, src:['asset','liab'] },
    { key:'income', label:'収入', badge:'bi', accts:getAccounts('income', true), sign:-1 },
    { key:'expense', label:'費用', badge:'bx', accts:getAccounts('expense', true), sign:1 },
    { key:'profit', label:'利益', badge:'bp', computed:true, formula:(i,x)=>i-x, src:['income','expense'] }
  ];

  const totals = {};
  const cols = '<col class="cl">' + months.map(() => '<col class="cm">').join('');
  const head = '<thead><tr><th>科目</th>' + months.map(m => `<th>${mLabel(m)}</th>`).join('') + '</tr></thead>';
  let body = '<tbody>';

  SCHEMA.forEach(g => {
    if (g.computed) {
      const s1 = totals[g.src[0]] || months.map(() => 0);
      const s2 = totals[g.src[1]] || months.map(() => 0);
      const tv = months.map((_,i) => g.formula(s1[i], s2[i]));
      totals[g.key] = tv;
      body += `<tr class="ghdr"><td><span class="badge ${g.badge}">${g.label}</span></td>${months.map(() => '<td></td>').join('')}</tr>`;
      body += `<tr class="sub"><td style="padding-left:16px;">${g.label}</td>${tv.map(v => `<td>${fmtSC(v)}</td>`).join('')}</tr>`;
      return;
    }

    body += `<tr class="ghdr"><td><span class="badge ${g.badge}">${g.label}</span></td>${months.map(() => '<td></td>').join('')}</tr>`;
    const at = months.map(() => 0);

    g.accts.forEach(a => {
      const vs = months.map(m => acctFlow(a, m) * g.sign);
      if (vs.every(v => v === 0)) return;
      vs.forEach((v,i) => at[i] += v);
      body += `<tr class="acct"><td>${escapeHtml(a)}</td>${vs.map(v => `<td>${fmtS(v)}</td>`).join('')}</tr>`;
    });

    totals[g.key] = at;
    body += `<tr class="sub"><td style="padding-left:16px;">${g.label}合計</td>${at.map(v => `<td>${fmtSC(v)}</td>`).join('')}</tr>`;
  });

  body += '</tbody>';
  document.getElementById('stbl').innerHTML = `<colgroup>${cols}</colgroup>${head}${body}`;
}

function buildFSMonthOptions() {
  const sel = document.getElementById('fs-month');
  const months = getMonths(12);
  sel.innerHTML = months.map((m, i) => {
    const y = m.getFullYear();
    const mo = m.getMonth();
    return `<option value="${y}-${mo}" ${i === months.length - 1 ? 'selected' : ''}>${y}年${mo + 1}月</option>`;
  }).join('');
}

function getSelectedMonth() {
  const v = document.getElementById('fs-month').value.split('-');
  return new Date(parseInt(v[0], 10), parseInt(v[1], 10), 1);
}

function renderBS() {
  const baseMonth = getSelectedMonth();

  function bsRows(accts, signFn) {
    let total = 0;
    const rows = [];
    accts.forEach(a => {
      const v = signFn(acctCumul(a, baseMonth));
      if (v !== 0) {
        rows.push({ name:a, val:v });
        total += v;
      }
    });
    return { rows, total };
  }

  const assetG = bsRows(getAccounts('asset', true), v => v);
  const liabG = bsRows(getAccounts('liability', true), v => -v);
  const equity = assetG.total - liabG.total;

  function colHtml(title, badge, sections, footLabel, footVal) {
    let h = `<div class="bs-col"><div class="bs-hdr"><span class="badge ${badge}">${title}</span></div>`;
    sections.forEach(s => {
      if (s.sectionLabel) h += `<div class="bs-row sec-hdr">${s.sectionLabel}</div>`;
      s.rows.forEach(r => {
        h += `<div class="bs-row acct"><span>${escapeHtml(r.name)}</span><span class="${r.val >= 0 ? 'pos' : 'neg'}">${fmt(r.val)}</span></div>`;
      });
      h += `<div class="bs-row sub"><span>${s.label}合計</span><span class="${s.total >= 0 ? 'pos' : 'neg'}">${fmt(s.total)}</span></div>`;
    });
    h += `<div class="bs-row sub" style="border-top:1px solid var(--border2);"><span>${footLabel}</span><span class="${footVal >= 0 ? 'pos' : 'neg'}">${fmt(footVal)}</span></div>`;
    return h + '</div>';
  }

  document.getElementById('bs-wrap').innerHTML =
    colHtml('資産', 'ba', [{label:'資産', ...assetG}], '資産合計', assetG.total) +
    colHtml('負債・純資産', 'be', [
      { sectionLabel:'負債', label:'負債', ...liabG },
      { sectionLabel:'純資産', label:'純資産', rows:[{ name:'純資産', val:equity }], total:equity }
    ], '負債＋純資産合計', liabG.total + equity);

  const months = getMonths(12);
  const eqData = months.map(m => {
    const a = getAccounts('asset', true).reduce((s,n) => s + acctCumul(n,m), 0);
    const l = getAccounts('liability', true).reduce((s,n) => s - acctCumul(n,m), 0);
    return a - l;
  });

  if (gBS) gBS.destroy();
  gBS = new Chart(document.getElementById('bs-chart'), {
    type:'line',
    data:{ labels:months.map(mLabel), datasets:[{ label:'純資産', data:eqData, borderColor:'#378ADD', backgroundColor:'rgba(55,138,221,0.08)', borderWidth:2, pointRadius:3, pointBackgroundColor:'#378ADD', tension:0.3, fill:true }] },
    options:{ responsive:true, maintainAspectRatio:false, plugins:{ legend:{ display:false } }, scales:{ x:{ grid:{ display:false }, ticks:{ color:'#888', font:{ size:10 }, autoSkip:false, maxRotation:0 } }, y:{ grid:{ color:'rgba(128,128,128,0.12)' }, ticks:{ color:'#888', font:{ size:9 }, callback:v=>'¥'+v.toLocaleString() } } } }
  });
}

function renderPL() {
  const baseMonth = getSelectedMonth();
  const es = mEntries(baseMonth);

  const incRows = [];
  let incTotal = 0;
  getAccounts('income', true).forEach(a => {
    const v = es.filter(e => e.crCat === a).reduce((s,e) => s + e.amount, 0);
    if (v > 0) { incRows.push({ name:a, val:v }); incTotal += v; }
  });

  const expRows = [];
  let expTotal = 0;
  getAccounts('expense', true).forEach(a => {
    const v = es.filter(e => e.drCat === a).reduce((s,e) => s + e.amount, 0);
    if (v > 0) { expRows.push({ name:a, val:v }); expTotal += v; }
  });

  const profit = incTotal - expTotal;
  let h = '';
  h += '<div class="pl-row hdr"><span class="badge bi">収入</span></div>';
  if (incRows.length) incRows.forEach(r => { h += `<div class="pl-row acct"><span>${escapeHtml(r.name)}</span><span class="pos">${fmt(r.val)}</span></div>`; });
  else h += '<div class="pl-row acct" style="color:var(--text2);">（なし）</div>';
  h += `<div class="pl-row sub"><span>収入合計</span><span class="pos">${fmt(incTotal)}</span></div>`;

  h += '<div class="pl-row hdr"><span class="badge bx">費用</span></div>';
  if (expRows.length) expRows.forEach(r => { h += `<div class="pl-row acct"><span>${escapeHtml(r.name)}</span><span class="neg">${fmt(r.val)}</span></div>`; });
  else h += '<div class="pl-row acct" style="color:var(--text2);">（なし）</div>';
  h += `<div class="pl-row sub"><span>費用合計</span><span class="neg">${fmt(expTotal)}</span></div>`;
  h += `<div class="pl-row total"><span><span class="badge bp">当月利益</span></span><span class="${profit >= 0 ? 'pos' : 'neg'}" style="font-size:15px;">${profit < 0 ? '-' : ''}${fmt(profit)}</span></div>`;

  document.getElementById('pl-wrap').innerHTML = h;

  const months = getMonths(12);
  const incD = months.map(m => mEntries(m).filter(isIncome).reduce((s,e) => s + e.amount, 0));
  const expD = months.map(m => mEntries(m).filter(isExpense).reduce((s,e) => s + e.amount, 0));
  const profD = months.map((_,i) => incD[i] - expD[i]);

  document.getElementById('pl-lgd').innerHTML =
    `<span><span class="ldot" style="background:#1D9E75;"></span>収入</span>
     <span><span class="ldot" style="background:#D85A30;"></span>費用</span>
     <span><span class="ldot" style="background:#378ADD;"></span>利益</span>`;

  if (gPL) gPL.destroy();
  gPL = new Chart(document.getElementById('pl-chart'), {
    type:'bar',
    data:{
      labels:months.map(mLabel),
      datasets:[
        { label:'収入', data:incD, backgroundColor:'rgba(29,158,117,0.6)', borderRadius:3, order:2 },
        { label:'費用', data:expD, backgroundColor:'rgba(216,90,48,0.6)', borderRadius:3, order:2 },
        { label:'利益', data:profD, type:'line', borderColor:'#378ADD', backgroundColor:'transparent', borderWidth:2, pointRadius:3, pointBackgroundColor:'#378ADD', tension:0.3, order:1 }
      ]
    },
    options:{ responsive:true, maintainAspectRatio:false, plugins:{ legend:{ display:false } }, scales:{ x:{ grid:{ display:false }, ticks:{ color:'#888', font:{ size:10 }, autoSkip:false, maxRotation:0 } }, y:{ grid:{ color:'rgba(128,128,128,0.12)' }, ticks:{ color:'#888', font:{ size:9 }, callback:v=>'¥'+v.toLocaleString() } } } }
  });
}

function buildMonthlyClosingSnapshot(baseMonth) {
  const es = mEntries(baseMonth);
  const assetBalances = {};
  const liabilityBalances = {};

  getAccounts('asset', true).forEach(name => {
    const value = acctCumul(name, baseMonth);
    if (value !== 0) assetBalances[name] = value;
  });

  getAccounts('liability', true).forEach(name => {
    const value = -acctCumul(name, baseMonth);
    if (value !== 0) liabilityBalances[name] = value;
  });

  const incomeTotal = es.filter(isIncome).reduce((s, e) => s + e.amount, 0);
  const expenseTotal = es.filter(isExpense).reduce((s, e) => s + e.amount, 0);
  const assetTotal = Object.values(assetBalances).reduce((s, v) => s + v, 0);
  const liabilityTotal = Object.values(liabilityBalances).reduce((s, v) => s + v, 0);

  return {
    month:monthKeyFromMonth(baseMonth),
    closedAt:new Date().toISOString(),
    assetBalances,
    liabilityBalances,
    incomeTotal,
    expenseTotal,
    profit:incomeTotal - expenseTotal,
    assetTotal,
    liabilityTotal,
    equity:assetTotal - liabilityTotal,
    entryCount:es.length
  };
}

function renderBalanceRows(title, balances) {
  const rows = Object.entries(balances);
  if (!rows.length) {
    return `<div class="closing-row muted"><span>${title}</span><span>データなし</span></div>`;
  }
  return rows.map(([name, value]) => `
    <div class="closing-row acct">
      <span>${escapeHtml(name)}</span>
      <span>${fmt(value)}</span>
    </div>
  `).join('');
}

function renderMonthlyClose() {
  const wrap = document.getElementById('monthly-close-wrap');
  if (!wrap) return;

  const baseMonth = getSelectedMonth();
  const snapshot = buildMonthlyClosingSnapshot(baseMonth);
  const closed = getClosing(snapshot.month);
  const shown = closed || snapshot;

  wrap.innerHTML = `
    <div class="closing-head">
      <div>
        <div class="closing-title">${escapeHtml(snapshot.month)} 月次締め</div>
        <div class="closing-sub">${closed ? `締め済み: ${escapeHtml(formatSyncTime(closed.closedAt))}` : '未締めです'}</div>
      </div>
      <span class="acct-tag ${closed ? 'on' : 'off'}">${closed ? '締め済み' : '未締め'}</span>
    </div>

    <div class="closing-grid">
      <div class="closing-metric"><span>収入</span><strong class="pos">${fmt(shown.incomeTotal)}</strong></div>
      <div class="closing-metric"><span>支出</span><strong class="neg">${fmt(shown.expenseTotal)}</strong></div>
      <div class="closing-metric"><span>収支</span><strong class="${shown.profit >= 0 ? 'pos' : 'neg'}">${fmtSigned(shown.profit)}</strong></div>
      <div class="closing-metric"><span>取引数</span><strong>${shown.entryCount}件</strong></div>
    </div>

    <div class="closing-columns">
      <div>
        <div class="closing-section-title">月末資産</div>
        ${renderBalanceRows('資産', shown.assetBalances)}
        <div class="closing-row total"><span>資産合計</span><span>${fmt(shown.assetTotal)}</span></div>
      </div>
      <div>
        <div class="closing-section-title">月末負債</div>
        ${renderBalanceRows('負債', shown.liabilityBalances)}
        <div class="closing-row total"><span>負債合計</span><span>${fmt(shown.liabilityTotal)}</span></div>
      </div>
    </div>
    <div class="closing-row total equity"><span>純資産</span><span>${fmt(shown.equity)}</span></div>

    <div class="sub-actions">
      ${closed
        ? '<button class="btnsub" type="button" onclick="reopenMonth()">締めを解除</button>'
        : '<button class="btnsub" type="button" onclick="closeMonth()">この月を締める</button>'
      }
    </div>
  `;
}

function closeMonth() {
  const baseMonth = getSelectedMonth();
  const snapshot = buildMonthlyClosingSnapshot(baseMonth);
  if (!confirm(`${snapshot.month} を締めますか？締め済み月の取引は、締め解除まで追加・編集・削除できません。`)) return;

  monthlyClosings = monthlyClosings.filter(c => c.month !== snapshot.month);
  monthlyClosings.push(snapshot);
  monthlyClosings.sort((a, b) => a.month.localeCompare(b.month));
  saveMonthlyClosings();
  renderMonthlyClose();
  alert('月次締めを実行しました');
}

function reopenMonth() {
  const month = monthKeyFromMonth(getSelectedMonth());
  if (!getClosing(month)) return;
  if (!confirm(`${month} の締めを解除しますか？`)) return;

  monthlyClosings = monthlyClosings.filter(c => c.month !== month);
  saveMonthlyClosings();
  renderMonthlyClose();
  alert('締めを解除しました');
}

function renderFS() {
  renderBS();
  renderPL();
  renderMonthlyClose();
}

function swFS(t) {
  document.querySelectorAll('.fs-tab').forEach((el, i) => el.classList.toggle('active', ['bs','pl'][i] === t));
  document.querySelectorAll('.fs-sec').forEach(el => el.classList.remove('active'));
  document.getElementById('fs-' + t).classList.add('active');
}

function sw(t) {
  const tabs = ['record','list','graph','budget','summary','fs','settings'];
  document.querySelectorAll('.tab-btn').forEach((el, i) => el.classList.toggle('active', tabs[i] === t));
  document.querySelectorAll('.sec').forEach(el => el.classList.remove('active'));
  document.getElementById('t-' + t).classList.add('active');

  if (t === 'list') renderList();
  if (t === 'graph') renderGraph();
  if (t === 'budget') renderBudget();
  if (t === 'summary') renderSummary();
  if (t === 'fs') {
    buildFSMonthOptions();
    renderFS();
  }
  if (t === 'settings') renderSettings();
}

function newEntryId() {
  return `${Date.now()}-${Math.random().toString(16).slice(2, 8)}`;
}

function isPointLinkedEntry(entry) {
  return entry?.preset === 'point' && entry.linkedId;
}

function isPointExpenseEntry(entry) {
  return isPointLinkedEntry(entry) && entry.pointRole === POINT_ROLE_EXPENSE;
}

function isPointAdjustmentEntry(entry) {
  return isPointLinkedEntry(entry) && entry.pointRole === POINT_ROLE_ADJUSTMENT;
}

function getLinkedPointEntries(entry) {
  if (!isPointLinkedEntry(entry)) return [entry].filter(Boolean);
  return entries.filter(e => e.linkedId === entry.linkedId);
}

function getPointExpenseEntry(entry) {
  if (!isPointLinkedEntry(entry)) return entry;
  return getLinkedPointEntries(entry).find(isPointExpenseEntry) || entry;
}

function makePointEntries(data, linkedId = newEntryId(), expenseId = newEntryId()) {
  const pointAsset = data.crCat || POINT_ASSET_ACCOUNT;
  const adjustmentDesc = data.desc ? `${data.desc}（ポイント発生）` : 'ポイント発生';
  const adjustmentNote = data.crNote || 'ポイント払いの自動補助';

  return [
    {
      id:newEntryId(),
      date:data.date,
      amount:data.amount,
      desc:adjustmentDesc,
      drCat:pointAsset,
      drNote:adjustmentNote,
      crCat:POINT_ADJUSTMENT_INCOME,
      crNote:adjustmentNote,
      preset:'point',
      entryType:'standard',
      linkedId,
      pointRole:POINT_ROLE_ADJUSTMENT
    },
    {
      ...data,
      id:expenseId,
      preset:'point',
      entryType:'standard',
      linkedId,
      pointRole:POINT_ROLE_EXPENSE
    }
  ];
}

function addEntry() {
  const date = document.getElementById('f-date').value;
  const amount = parseFloat(document.getElementById('f-amt').value);
  const desc = document.getElementById('f-desc').value.trim();
  const eventId = currentPreset === 'opening' ? undefined : (document.getElementById('f-event').value || undefined);
  const drCat = document.getElementById('f-dr').value;
  const drNote = document.getElementById('f-dr-note').value.trim();
  const crCat = document.getElementById('f-cr').value;
  const crNote = document.getElementById('f-cr-note').value.trim();

  if (!date || !amount || amount <= 0) {
    alert('日付と金額を入力してください');
    return;
  }
  if (!drCat || !crCat) {
    alert('勘定科目を選択してください');
    return;
  }
  if (eventId) {
    const event = getBudgetEvent(eventId);
    if (event && date && !isEventDateMatch(event, date)) {
      if (!confirm(`「${event.name}」の期間は ${event.startDate || '未設定'} - ${event.endDate || '未設定'} です。期間外の支出として登録しますか？`)) {
        return;
      }
    }
  }

  const targetMonth = monthKeyFromDate(date);
  if (blockIfClosedMonth(targetMonth, editingId ? '更新' : '追加')) return;

  const entryType = currentPreset === 'opening' ? 'opening' : 'standard';
  const data = { date, amount, desc, eventId, drCat, drNote, crCat, crNote, preset:currentPreset, entryType };

  if (currentPreset === 'opening' && !isValidOpeningEntry(data)) {
    alert(`開始残高は「資産 / ${OPENING_BALANCE_EQUITY}」または「${OPENING_BALANCE_EQUITY} / 負債」の形で入力してください。`);
    return;
  }

  uiPrefs.lastPreset = currentPreset;
  uiPrefs.lastCreditByPreset = uiPrefs.lastCreditByPreset || {};
  uiPrefs.lastDebitByPreset = uiPrefs.lastDebitByPreset || {};
  uiPrefs.lastCreditByPreset[currentPreset] = crCat;
  uiPrefs.lastDebitByPreset[currentPreset] = drCat;
  saveUiPrefs();

  if (editingId) {
    const idx = entries.findIndex(e => e.id === editingId);
    if (idx === -1) {
      alert('編集対象が見つかりませんでした');
      cancelEdit(false);
      return;
    }
    const original = entries[idx];
    const originalLinked = getLinkedPointEntries(original);
    if (originalLinked.some(e => blockIfClosedMonth(monthKeyFromDate(e.date), '更新'))) return;

    if (currentPreset === 'point') {
      const linkedId = isPointLinkedEntry(original) ? original.linkedId : newEntryId();
      const expenseId = isPointLinkedEntry(original)
        ? getPointExpenseEntry(original).id
        : original.id;
      const pointEntries = makePointEntries(data, linkedId, expenseId);
      const linkedIds = new Set(originalLinked.map(e => e.id));
      entries = entries.filter(e => !linkedIds.has(e.id));
      entries.push(...pointEntries);
    } else {
      const linkedIds = new Set(originalLinked.map(e => e.id));
      entries = entries.filter(e => !linkedIds.has(e.id));
      entries.push({ ...original, ...data, preset:currentPreset, entryType, linkedId:undefined, pointRole:undefined });
    }
    saveEntries();
    cancelEdit(false);
    refreshActiveTab();
    renderQuickCreditButtons();
    alert('更新しました！');
    return;
  }

  if (currentPreset === 'point') {
    entries.push(...makePointEntries(data));
  } else {
    entries.push({ id:newEntryId(), ...data });
  }
  saveEntries();
  refreshActiveTab();
  resetForm();
  applyLastUsedForm();
  renderQuickCreditButtons();
  moveFocusAfterAdd();
  alert('追加しました！');
}

function startEdit(id) {
  const clickedEntry = entries.find(e => e.id === id);
  const entry = getPointExpenseEntry(clickedEntry);
  if (!entry) {
    alert('編集対象が見つかりませんでした');
    return;
  }
  if (getLinkedPointEntries(entry).some(e => blockIfClosedMonth(monthKeyFromDate(e.date), '編集'))) return;

  editingId = entry.id;
  setPreset(entry.preset || guessPreset(entry));

  document.getElementById('f-date').value = entry.date || '';
  document.getElementById('f-amt').value = entry.amount ?? '';
  document.getElementById('f-desc').value = entry.desc || '';
  renderEventOptions(entry.eventId || '');
  document.getElementById('f-event').value = entry.eventId || '';
  document.getElementById('f-dr').value = entry.drCat || '';
  document.getElementById('f-dr-note').value = entry.drNote || '';
  document.getElementById('f-cr').value = entry.crCat || '';
  document.getElementById('f-cr-note').value = entry.crNote || '';

  renderQuickCreditButtons();

  document.getElementById('submit-btn').textContent = '更新';
  const editBar = document.getElementById('edit-bar');
  editBar.classList.add('show');
  const pairNote = isPointLinkedEntry(entry) ? '<br>ポイント発生の補助仕訳も一緒に更新します。' : '';
  editBar.innerHTML = `編集中です。<br>${escapeHtml(entry.desc || entry.drCat)} / ${escapeHtml(entry.date)} / ${fmt(entry.amount)}${pairNote}`;

  sw('record');
  window.scrollTo({ top:0, behavior:'smooth' });
}

function cancelEdit(showAlert = false) {
  editingId = null;
  document.getElementById('submit-btn').textContent = '追加';
  document.getElementById('edit-bar').classList.remove('show');
  document.getElementById('edit-bar').textContent = '';
  resetForm();
  if (showAlert) alert('編集をキャンセルしました');
}

function resetForm() {
  document.getElementById('f-date').value = formatLocalDate(new Date());
  document.getElementById('f-amt').value = '';
  document.getElementById('f-desc').value = '';
  renderEventOptions('');
  document.getElementById('f-dr-note').value = '';
  document.getElementById('f-cr-note').value = '';
  setPreset(uiPrefs.lastPreset || 'expense');
  renderQuickCreditButtons();
}

function guessPreset(entry) {
  const assets = getAccounts('asset', true);
  const liabilities = getAccounts('liability', true);
  const incomes = getAccounts('income', true);

  if (isOpeningEntry(entry) || entry.crCat === OPENING_BALANCE_EQUITY || entry.drCat === OPENING_BALANCE_EQUITY) return 'opening';
  if (entry.preset === 'point' || entry.linkedId && entry.pointRole) return 'point';
  if (incomes.includes(entry.crCat) && assets.includes(entry.drCat)) return 'income';
  if (liabilities.includes(entry.drCat) && assets.includes(entry.crCat)) return 'repay';
  if (assets.includes(entry.drCat) && assets.includes(entry.crCat)) return 'transfer';
  return 'expense';
}

function buildBackupPayload() {
  return {
    app:'kakeibo',
    version:7,
    exportedAt:new Date().toISOString(),
    entries,
    accountSettings,
    uiPrefs,
    monthlyClosings,
    monthlyBudgets,
    budgetEvents
  };
}

function applyBackupPayload(raw) {
  const importedEntries = Array.isArray(raw) ? raw : raw.entries;
  if (!Array.isArray(importedEntries)) throw new Error('バックアップ形式が不正です');

  const normalized = importedEntries.map(normalizeEntry);
  if (normalized.some(v => !v)) throw new Error('読み込めないデータが含まれています');

  entries = normalized;
  if (raw.accountSettings) {
    accountSettings = {
      asset: normalizeAccountBlock(raw.accountSettings.asset, 'asset'),
      liability: normalizeAccountBlock(raw.accountSettings.liability, 'liability'),
      income: normalizeAccountBlock(raw.accountSettings.income, 'income'),
      expense: normalizeAccountBlock(raw.accountSettings.expense, 'expense')
    };
    saveAccountSettings();
  }

  if (raw.uiPrefs && typeof raw.uiPrefs === 'object') {
    uiPrefs = raw.uiPrefs;
    saveUiPrefs();
  }

  monthlyClosings = Array.isArray(raw.monthlyClosings)
    ? raw.monthlyClosings.map(normalizeMonthlyClosing).filter(Boolean)
    : [];
  localStorage.setItem('kakeibo_monthly_closings', JSON.stringify(monthlyClosings));

  monthlyBudgets = raw.monthlyBudgets && typeof raw.monthlyBudgets === 'object'
    ? Object.fromEntries(
        Object.entries(raw.monthlyBudgets)
          .filter(([month]) => /^\d{4}-\d{2}$/.test(month))
          .map(([month, value]) => [month, normalizeMonthlyBudget(value)])
      )
    : {};
  localStorage.setItem('kakeibo_monthly_budgets', JSON.stringify(monthlyBudgets));

  budgetEvents = Array.isArray(raw.budgetEvents)
    ? raw.budgetEvents.map(normalizeBudgetEvent).filter(Boolean)
    : [];
  localStorage.setItem('kakeibo_budget_events', JSON.stringify(budgetEvents));

  saveEntries();
  cancelEdit(false);
  cancelBudgetEventEdit(false);
  refreshActiveTab();
  refreshAccountDrivenUI();
}

function exportData() {
  const payload = buildBackupPayload();

  const blob = new Blob([JSON.stringify(payload, null, 2)], { type:'application/json' });
  const url = URL.createObjectURL(blob);

  const d = new Date();
  const pad = n => String(n).padStart(2, '0');
  const filename = `kakeibo-backup-${d.getFullYear()}${pad(d.getMonth()+1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}.json`;

  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function importData(event) {
  const file = event.target.files && event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const raw = JSON.parse(e.target.result);
      const importedEntries = Array.isArray(raw) ? raw : raw.entries;
      if (!Array.isArray(importedEntries)) throw new Error('バックアップ形式が不正です');

      if (!confirm(`現在のデータを消して、${importedEntries.length}件のバックアップで置き換えますか？`)) {
        event.target.value = '';
        return;
      }

      applyBackupPayload(raw);
      alert('バックアップを読み込みました');
    } catch (err) {
      alert('読み込みに失敗しました: ' + err.message);
    } finally {
      event.target.value = '';
    }
  };
  reader.readAsText(file);
}

function loadSyncSettings() {
  const saved = JSON.parse(localStorage.getItem('kakeibo_sync_settings') || '{}');
  return {
    endpoint:saved.endpoint || '',
    token:saved.token || '',
    lastSyncAt:saved.lastSyncAt || '',
    lastLocalChangeAt:saved.lastLocalChangeAt || ''
  };
}

function saveSyncSettings() {
  localStorage.setItem('kakeibo_sync_settings', JSON.stringify(syncSettings));
}

function markLocalChanged() {
  if (!syncSettings) return;
  syncSettings.lastLocalChangeAt = new Date().toISOString();
  saveSyncSettings();
  renderSyncStatus();
}

function saveSyncConfig(message = '同期設定を保存しました') {
  syncSettings.endpoint = (document.getElementById('sync-endpoint')?.value || '').trim();
  syncSettings.token = (document.getElementById('sync-token')?.value || '').trim();
  const passphrase = (document.getElementById('sync-passphrase')?.value || '').trim();
  if (passphrase) sessionSyncPassphrase = passphrase;
  saveSyncSettings();
  renderSyncStatus(message);
}

function getSyncHeaders() {
  const headers = { 'Content-Type':'application/json' };
  if (syncSettings.token) headers.Authorization = `Bearer ${syncSettings.token}`;
  return headers;
}

function ensureSyncEndpoint() {
  saveSyncConfig('');
  if (!syncSettings.endpoint) {
    throw new Error('Cloudflare Worker URLを入力してください');
  }
}

function getSyncPassphrase() {
  const input = document.getElementById('sync-passphrase');
  const typed = (input?.value || '').trim();
  if (typed) {
    sessionSyncPassphrase = typed;
    if (input) input.value = '';
    renderSyncStatus('暗号化パスフレーズをこのタブで記憶しました');
    return sessionSyncPassphrase;
  }

  if (sessionSyncPassphrase) return sessionSyncPassphrase;

  const prompted = prompt('暗号化パスフレーズを入力してください。パスフレーズはこの端末に保存されません。');
  if (!prompted) throw new Error('暗号化パスフレーズが必要です');
  sessionSyncPassphrase = prompted;
  renderSyncStatus('暗号化パスフレーズをこのタブで記憶しました');
  return sessionSyncPassphrase;
}

function bytesToBase64(bytes) {
  let binary = '';
  bytes.forEach(b => { binary += String.fromCharCode(b); });
  return btoa(binary);
}

function base64ToBytes(value) {
  const binary = atob(value);
  return Uint8Array.from(binary, ch => ch.charCodeAt(0));
}

async function deriveCryptoKey(passphrase, salt) {
  if (!window.crypto?.subtle) throw new Error('このブラウザは暗号化に対応していません');
  const material = await crypto.subtle.importKey(
    'raw',
    new TextEncoder().encode(passphrase),
    'PBKDF2',
    false,
    ['deriveKey']
  );
  return crypto.subtle.deriveKey(
    { name:'PBKDF2', salt, iterations:210000, hash:'SHA-256' },
    material,
    { name:'AES-GCM', length:256 },
    false,
    ['encrypt', 'decrypt']
  );
}

async function encryptPayload(payload, passphrase) {
  const salt = crypto.getRandomValues(new Uint8Array(16));
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const key = await deriveCryptoKey(passphrase, salt);
  const plaintext = new TextEncoder().encode(JSON.stringify(payload));
  const ciphertext = await crypto.subtle.encrypt({ name:'AES-GCM', iv }, key, plaintext);

  return {
    app:'kakeibo',
    version:1,
    encrypted:true,
    crypto:{
      algorithm:'AES-GCM',
      kdf:'PBKDF2',
      hash:'SHA-256',
      iterations:210000
    },
    salt:bytesToBase64(salt),
    iv:bytesToBase64(iv),
    ciphertext:bytesToBase64(new Uint8Array(ciphertext)),
    savedAt:new Date().toISOString()
  };
}

async function decryptPayload(envelope, passphrase) {
  if (!envelope?.encrypted || envelope.app !== 'kakeibo') {
    throw new Error('サーバーのデータ形式が不正です');
  }
  if (envelope.crypto?.algorithm !== 'AES-GCM' || envelope.crypto?.kdf !== 'PBKDF2') {
    throw new Error('対応していない暗号化形式です');
  }

  const salt = base64ToBytes(envelope.salt);
  const iv = base64ToBytes(envelope.iv);
  const ciphertext = base64ToBytes(envelope.ciphertext);
  const key = await deriveCryptoKey(passphrase, salt);

  try {
    const plaintext = await crypto.subtle.decrypt({ name:'AES-GCM', iv }, key, ciphertext);
    return JSON.parse(new TextDecoder().decode(plaintext));
  } catch (err) {
    throw new Error('復号に失敗しました。パスフレーズが違う可能性があります。');
  }
}

async function uploadToServer() {
  try {
    ensureSyncEndpoint();
    const passphrase = getSyncPassphrase();
    setSyncBusy(true, '暗号化してアップロード中...');
    const encryptedPayload = await encryptPayload(buildBackupPayload(), passphrase);

    const res = await fetch(syncSettings.endpoint, {
      method:'PUT',
      headers:getSyncHeaders(),
      body:JSON.stringify(encryptedPayload)
    });

    if (!res.ok) throw new Error(`サーバーが ${res.status} を返しました`);

    syncSettings.lastSyncAt = new Date().toISOString();
    saveSyncSettings();
    renderSyncStatus('暗号化してCloudflareに保存しました');
  } catch (err) {
    renderSyncStatus('アップロード失敗: ' + err.message, true);
  } finally {
    setSyncBusy(false);
  }
}

async function downloadFromServer() {
  try {
    ensureSyncEndpoint();
    if (!confirm('サーバーのデータで、この端末のデータを置き換えますか？')) return;
    const passphrase = getSyncPassphrase();

    setSyncBusy(true, 'ダウンロードして復号中...');
    const res = await fetch(syncSettings.endpoint, {
      method:'GET',
      headers:getSyncHeaders()
    });

    if (!res.ok) throw new Error(`サーバーが ${res.status} を返しました`);

    const raw = await res.json();
    const payload = await decryptPayload(raw, passphrase);
    const count = Array.isArray(payload) ? payload.length : payload.entries.length;
    applyBackupPayload(payload);

    syncSettings.lastSyncAt = new Date().toISOString();
    saveSyncSettings();
    renderSyncStatus(`Cloudflareから${count}件を復元しました`);
    alert('Cloudflareのデータを読み込みました');
  } catch (err) {
    renderSyncStatus('ダウンロード失敗: ' + err.message, true);
  } finally {
    setSyncBusy(false);
  }
}

function setSyncBusy(isBusy, message = '') {
  document.querySelectorAll('[data-sync-action]').forEach(btn => {
    btn.disabled = isBusy;
  });
  if (message) renderSyncStatus(message);
}

function formatSyncTime(value) {
  if (!value) return '未実行';
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return '未実行';
  return d.toLocaleString('ja-JP', { dateStyle:'short', timeStyle:'short' });
}

function renderSyncStatus(message = '', isError = false) {
  const status = document.getElementById('sync-status');
  if (!status) return;

  const lastSync = formatSyncTime(syncSettings.lastSyncAt);
  const lastChange = formatSyncTime(syncSettings.lastLocalChangeAt);
  const passState = sessionSyncPassphrase ? 'このタブで記憶中' : '未入力';
  status.classList.toggle('err', isError);
  status.innerHTML = `
    <div>${message ? escapeHtml(message) : 'Cloudflareへ暗号化して手動保存・復元できます。'}</div>
    <div class="sync-meta">最終同期: ${escapeHtml(lastSync)} / この端末の最終変更: ${escapeHtml(lastChange)} / パスフレーズ: ${passState}</div>
  `;
}

function normalizeEntry(e) {
  if (!e || typeof e !== 'object') return null;
  if (!e.date || !e.drCat || !e.crCat) return null;
  const amount = Number(e.amount);
  if (!Number.isFinite(amount) || amount <= 0) return null;

  return {
    id: String(e.id || `${Date.now()}-${Math.random().toString(16).slice(2, 8)}`),
    date: String(e.date),
    amount,
    desc: String(e.desc || ''),
    eventId: e.eventId ? String(e.eventId) : undefined,
    drCat: String(e.drCat),
    drNote: String(e.drNote || ''),
    crCat: String(e.crCat),
    crNote: String(e.crNote || ''),
    preset: String(e.preset || guessPreset(e)),
    entryType: isOpeningEntry(e) ? 'opening' : 'standard',
    linkedId: e.linkedId ? String(e.linkedId) : undefined,
    pointRole: e.pointRole ? String(e.pointRole) : undefined
  };
}

function normalizeNumberMap(value) {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return {};
  return Object.fromEntries(
    Object.entries(value)
      .map(([key, val]) => [String(key), Number(val)])
      .filter(([, val]) => Number.isFinite(val))
  );
}

function normalizeMonthlyClosing(closing) {
  if (!closing || typeof closing !== 'object') return null;
  const month = String(closing.month || '');
  if (!/^\d{4}-\d{2}$/.test(month)) return null;

  const assetBalances = normalizeNumberMap(closing.assetBalances);
  const liabilityBalances = normalizeNumberMap(closing.liabilityBalances);
  const incomeTotal = Number(closing.incomeTotal || 0);
  const expenseTotal = Number(closing.expenseTotal || 0);
  const assetTotal = Number(closing.assetTotal ?? Object.values(assetBalances).reduce((s, v) => s + v, 0));
  const liabilityTotal = Number(closing.liabilityTotal ?? Object.values(liabilityBalances).reduce((s, v) => s + v, 0));

  return {
    month,
    closedAt:String(closing.closedAt || new Date().toISOString()),
    assetBalances,
    liabilityBalances,
    incomeTotal:Number.isFinite(incomeTotal) ? incomeTotal : 0,
    expenseTotal:Number.isFinite(expenseTotal) ? expenseTotal : 0,
    profit:Number.isFinite(Number(closing.profit)) ? Number(closing.profit) : incomeTotal - expenseTotal,
    assetTotal:Number.isFinite(assetTotal) ? assetTotal : 0,
    liabilityTotal:Number.isFinite(liabilityTotal) ? liabilityTotal : 0,
    equity:Number.isFinite(Number(closing.equity)) ? Number(closing.equity) : assetTotal - liabilityTotal,
    entryCount:Number.isFinite(Number(closing.entryCount)) ? Number(closing.entryCount) : 0
  };
}

function normalizeMonthlyBudget(value) {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return {};
  return Object.fromEntries(
    Object.entries(value)
      .map(([category, amount]) => [String(category), Number(amount)])
      .filter(([, amount]) => Number.isFinite(amount) && amount > 0)
      .map(([category, amount]) => [category, Math.round(amount)])
  );
}

function normalizeBudgetEvent(event) {
  if (!event || typeof event !== 'object') return null;
  const name = String(event.name || '').trim();
  const budget = Number(event.budget ?? event.amount ?? 0);
  if (!name || !Number.isFinite(budget) || budget < 0) return null;

  return {
    id:String(event.id || newEntryId()),
    name,
    budget:Math.round(budget),
    startDate:String(event.startDate || ''),
    endDate:String(event.endDate || '')
  };
}

function renderSettings(){
  const endpointInput = document.getElementById('sync-endpoint');
  const tokenInput = document.getElementById('sync-token');
  if (endpointInput && endpointInput.value !== syncSettings.endpoint) {
    endpointInput.value = syncSettings.endpoint;
  }
  if (tokenInput && tokenInput.value !== syncSettings.token) {
    tokenInput.value = syncSettings.token;
  }
  renderSyncStatus();

  ['asset','liability','income','expense'].forEach(type => {
    const mount = document.getElementById(`acct-list-${type}`);
    if (!mount) return;

    const list = accountSettings[type] || [];
    if (!list.length) {
      mount.innerHTML = '<div class="acct-empty">科目がありません</div>';
      return;
    }

    mount.innerHTML = list.map(item => {
      const used = isAccountUsed(type, item.name);
      const colorControl = type === 'expense'
        ? `
          <span class="color-dot" style="background:${item.color || '#888'};"></span>
          <input
            class="color-picker"
            type="color"
            value="${item.color || '#888888'}"
            onchange="updateExpenseColor('${escapeJs(item.name)}', this.value)"
            aria-label="${escapeHtml(item.name)}の色"
          >
          <select onchange="updateExpenseCostType('${escapeJs(item.name)}', this.value)">
            <option value="fixed" ${item.costType === 'fixed' ? 'selected' : ''}>固定費</option>
            <option value="variable" ${item.costType === 'fixed' ? '' : 'selected'}>変動費</option>
          </select>
        `
        : '';

      return `
        <div class="acct-row">
          <div class="acct-name-wrap">
            ${colorControl}
            <span class="acct-name">${escapeHtml(item.name)}</span>
            <span class="acct-tag ${item.active ? 'on' : 'off'}">${item.active ? '有効' : '無効'}</span>
            ${used ? '<span class="acct-meta">使用済み</span>' : '<span class="acct-meta">未使用</span>'}
          </div>
          <div class="acct-actions">
            ${item.active
              ? `<button class="acct-btn warn" onclick="disableOrDeleteAccount('${type}', '${escapeJs(item.name)}')">${used ? '無効化' : '削除'}</button>`
              : `<button class="acct-btn" onclick="enableAccount('${type}', '${escapeJs(item.name)}')">再有効化</button>`
            }
          </div>
        </div>
      `;
    }).join('');
  });
}

function addAccount(type){
  const input = document.getElementById(`acct-input-${type}`);
  const name = (input.value || '').trim();
  if (!name) {
    alert('科目名を入力してください');
    return;
  }

  if (hasAccount(type, name)) {
    const target = accountSettings[type].find(a => a.name === name);
    if (target && !target.active) {
      target.active = true;
      saveAccountSettings();
      renderSettings();
      refreshAccountDrivenUI();
      input.value = '';
      alert('無効化されていた科目を再有効化しました');
      return;
    }
    alert('その科目はすでに存在します');
    return;
  }

  if (type === 'expense') {
    accountSettings[type].push({ name, active:true, color:randomColor(), costType:'variable' });
  } else {
    accountSettings[type].push({ name, active:true });
  }

  saveAccountSettings();
  renderSettings();
  refreshAccountDrivenUI();
  input.value = '';
}

function disableOrDeleteAccount(type, name){
  const idx = accountSettings[type].findIndex(a => a.name === name);
  if (idx === -1) return;

  const used = isAccountUsed(type, name);
  if (used) {
    if (!confirm(`「${name}」は過去の記録で使われています。無効化しますか？`)) return;
    accountSettings[type][idx].active = false;
  } else {
    if (!confirm(`「${name}」を削除しますか？`)) return;
    accountSettings[type].splice(idx, 1);
  }

  saveAccountSettings();
  renderSettings();
  refreshAccountDrivenUI();
}

function enableAccount(type, name){
  const target = accountSettings[type].find(a => a.name === name);
  if (!target) return;
  target.active = true;
  saveAccountSettings();
  renderSettings();
  refreshAccountDrivenUI();
}

function updateExpenseColor(name, color){
  const target = accountSettings.expense.find(a => a.name === name);
  if (!target) return;
  target.color = color;
  saveAccountSettings();
  renderSettings();

  const active = getActiveTab();
  if (active === 'graph') renderGraph();
}

function updateExpenseCostType(name, costType) {
  const target = accountSettings.expense.find(a => a.name === name);
  if (!target) return;
  target.costType = costType === 'fixed' ? 'fixed' : 'variable';
  saveAccountSettings();
  renderSettings();
  refreshActiveTab();
}

function escapeHtml(str) {
  return String(str).replace(/[&<>"']/g, s => ({
    '&':'&amp;',
    '<':'&lt;',
    '>':'&gt;',
    '"':'&quot;',
    "'":'&#39;'
  }[s]));
}

function escapeJs(str){
  return String(str).replace(/\\/g, '\\\\').replace(/'/g, "\\'");
}

setPreset(uiPrefs.lastPreset || 'expense');
document.getElementById('f-date').value = formatLocalDate(new Date());
updateMetrics();
renderSettings();
renderEventOptions('');
renderBudget();
updateListCategoryFilterOptions();
renderQuickCreditButtons();
setupFormInteractions();
