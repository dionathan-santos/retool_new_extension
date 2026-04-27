// ─────────────────────────────────────────────────────────────────────────────
//  Retool Suite Importer — standalone console script
//
//  HOW TO USE
//  1. Open the Retool page in Chrome/Edge/Firefox.
//  2. Open DevTools (F12) → Console tab.
//  3. Paste this entire file and press Enter.
//  4. A file picker will appear — choose your .xlsx file.
//  5. The script logs progress to the console and fills each suite.
//
//  SETTINGS — edit the block below before pasting if you need different values.
// ─────────────────────────────────────────────────────────────────────────────

;(async function () {

  const SETTINGS = {
    delay:            2500,  // ms to wait between suites
    skipErrors:       false, // true = log suite errors and keep going
    submitAfterFill:  true,  // true = click Confirm/Save after each suite
    dryRun:           false, // true = fill fields but do NOT click Confirm/Save
    reviewBeforeSave: true,  // true = show floating panel so user can review before saving
  }

  // ─── Checkpoint / Resume ──────────────────────────────────────────────────
  // Saves import progress to localStorage after every suite so the session
  // can be resumed after a page refresh or accidental close.

  const SESSION_KEY = 'AY_IMPORT_v1'

  function sessionSave (data) {
    try { localStorage.setItem(SESSION_KEY, JSON.stringify(data)) } catch (_) {}
  }

  function sessionLoad () {
    try { return JSON.parse(localStorage.getItem(SESSION_KEY) ?? 'null') } catch (_) { return null }
  }

  function sessionClear () {
    try { localStorage.removeItem(SESSION_KEY) } catch (_) {}
  }

  // ─── Floating progress overlay ────────────────────────────────────────────
  // A small panel fixed to the bottom-right corner so the user can track
  // progress without keeping DevTools open.

  let _overlay = null

  function overlayCreate (total) {
    if (_overlay) _overlay.remove()
    const el = document.createElement('div')
    el.id = '__ay_overlay__'
    el.style.cssText = [
      'position:fixed', 'bottom:20px', 'right:20px', 'z-index:2147483647',
      'background:#1e1e2e', 'color:#cdd6f4', 'font:13px/1.5 monospace',
      'padding:14px 18px', 'border-radius:10px', 'min-width:260px',
      'box-shadow:0 4px 24px rgba(0,0,0,.5)', 'pointer-events:none',
    ].join(';')
    el.innerHTML = `
      <div style="font-weight:700;margin-bottom:6px;color:#89b4fa">
        ⚙ Retool Importer
      </div>
      <div id="__ay_bar_wrap__" style="background:#313244;border-radius:4px;height:6px;margin-bottom:8px">
        <div id="__ay_bar__" style="background:#a6e3a1;height:6px;border-radius:4px;width:0%;transition:width .3s"></div>
      </div>
      <div id="__ay_status__">Starting…</div>
      <div id="__ay_last__" style="color:#6c7086;font-size:11px;margin-top:4px"></div>
    `
    document.body.appendChild(el)
    _overlay = el
  }

  function overlayUpdate ({ done, total, label, status }) {
    if (!_overlay) return
    const pct = total ? Math.round((done / total) * 100) : 0
    const bar  = document.getElementById('__ay_bar__')
    const stat = document.getElementById('__ay_status__')
    const last = document.getElementById('__ay_last__')
    if (bar)  bar.style.width  = pct + '%'
    if (stat) stat.textContent = `Suite ${done}/${total} — ${label}`
    if (last && status) {
      const colour = status === 'success' ? '#a6e3a1' : status === 'error' ? '#f38ba8' : '#f9e2af'
      last.innerHTML = `<span style="color:${colour}">${status === 'success' ? '✓' : status === 'error' ? '✗' : '⏸'}</span> ${label}`
    }
  }

  function overlayDestroy (msg = 'Done') {
    if (!_overlay) return
    const stat = document.getElementById('__ay_status__')
    if (stat) stat.textContent = msg
    setTimeout(() => { _overlay?.remove(); _overlay = null }, 8000)
  }

  // ─── 1. Load SheetJS from CDN if not already present ─────────────────────

  async function loadSheetJS () {
    if (window.XLSX) return
    await new Promise((resolve, reject) => {
      const s = document.createElement('script')
      s.src = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js'
      s.onload  = resolve
      s.onerror = () => reject(new Error('Failed to load SheetJS from CDN'))
      document.head.appendChild(s)
    })
  }

  // ─── 2. Prompt user for file, return parsed ArrayBuffer ──────────────────

  function pickFile () {
    return new Promise((resolve, reject) => {
      const input = document.createElement('input')
      input.type   = 'file'
      input.accept = '.xlsx,.xls,.csv'
      input.style.display = 'none'
      document.body.appendChild(input)

      input.addEventListener('change', function () {
        const file = this.files[0]
        if (!file) { reject(new Error('No file selected')); return }
        const reader = new FileReader()
        reader.onload  = e => { document.body.removeChild(input); resolve({ name: file.name, buffer: e.target.result }) }
        reader.onerror = () => reject(new Error('FileReader error'))
        reader.readAsArrayBuffer(file)
      })

      // Some browsers require the input to be in the DOM before .click()
      setTimeout(() => input.click(), 0)
    })
  }

  // ─── 3. Parse workbook — same logic as popup.js ───────────────────────────

  function parseWorkbook (wb) {
    const sheetName = wb.SheetNames[0]
    const ws        = wb.Sheets[sheetName]
    const raw       = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' })

    const headerRowIndex = raw.findIndex(row =>
      row.some(cell => String(cell).trim() === 'Suite #')
    )
    if (headerRowIndex === -1) throw new Error('Header row with "Suite #" not found')

    const typeRow   = headerRowIndex > 0 ? (raw[headerRowIndex - 1] ?? []) : []
    const headerRow = raw[headerRowIndex]
    const dataRows  = raw.slice(headerRowIndex + 2)

    const suites = []
    for (const row of dataRows) {
      const suite = {}
      for (let c = 0; c < headerRow.length; c++) {
        const header = String(headerRow[c]).trim()
        if (!header) continue
        if (String(typeRow[c] ?? '').trim().toUpperCase() === 'SKIP') continue
        suite[header] = row[c] ?? ''
      }
      const suiteNum = String(suite['Suite #'] ?? '').trim()
      const address  = String(suite['Address']  ?? '').trim()
      if (!suiteNum && !address) continue
      suites.push(suite)
    }
    return suites
  }

  // ─── 4. Utilities ─────────────────────────────────────────────────────────

  function sleep (ms) { return new Promise(r => setTimeout(r, ms)) }

  function scrollIntoView (el) {
    try { el.scrollIntoView({ block: 'center', behavior: 'smooth' }) } catch (_) {}
  }

  // ─── 5. React / DOM helpers ───────────────────────────────────────────────

  function getReactFiber (el) {
    const key = Object.keys(el).find(k => k.startsWith('__reactFiber'))
    return key ? el[key] : null
  }

  function callReactProp (el, propName, ...args) {
    let fiber = getReactFiber(el)
    while (fiber) {
      if (fiber.memoizedProps?.[propName]) {
        fiber.memoizedProps[propName](...args)
        return true
      }
      fiber = fiber.return
    }
    return false
  }

  function getSelectFiberProps (el) {
    let fiber = getReactFiber(el)
    while (fiber) {
      const p = fiber.memoizedProps
      if (p?.onChange && p?.options) return p
      fiber = fiber.return
    }
    return null
  }

  const nativeInputSetter = Object.getOwnPropertyDescriptor(
    window.HTMLInputElement.prototype, 'value'
  ).set

  const nativeTextAreaSetter = Object.getOwnPropertyDescriptor(
    window.HTMLTextAreaElement.prototype, 'value'
  ).set

  function reactSet (el, value) {
    nativeInputSetter.call(el, value)
    el.dispatchEvent(new Event('input',  { bubbles: true }))
    el.dispatchEvent(new Event('change', { bubbles: true }))
  }

  async function focusSetBlur (el, value) {
    el.focus()
    await sleep(100)
    reactSet(el, value)
    await sleep(150)
    el.blur()
    el.dispatchEvent(new Event('focusout', { bubbles: true }))
    await sleep(100)
  }

  function reactSetTextArea (el, value) {
    nativeTextAreaSetter.call(el, value)
    el.dispatchEvent(new Event('input',  { bubbles: true }))
    el.dispatchEvent(new Event('change', { bubbles: true }))
  }

  // ─── 6. Date helpers ──────────────────────────────────────────────────────

  function excelDateToISO (serial) {
    const d = new Date((serial - 25569) * 86400 * 1000)
    return d.toISOString().slice(0, 10)
  }

  function normalizeDate (value) {
    if (value === null || value === undefined || value === '') return ''
    if (typeof value === 'number') return excelDateToISO(value)
    const str = String(value).trim()
    if (/^\d+$/.test(str)) {
      const n = parseInt(str, 10)
      if (n > 20000 && n < 80000) return excelDateToISO(n)
    }
    const d = new Date(str)
    if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10)
    return str
  }

  const MONTH_ABBR = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']

  function isoToDisplayDate (iso) {
    const [year, month, day] = iso.split('-').map(Number)
    return `${MONTH_ABBR[month - 1]} ${day}, ${year}`
  }

  // ─── 7. Per-type fill functions ───────────────────────────────────────────

  async function fillText (el, value) {
    scrollIntoView(el)
    await focusSetBlur(el, String(value))
  }

  async function fillNumber (el, value) {
    scrollIntoView(el)
    await focusSetBlur(el, String(value))
  }

  async function fillDate (el, value) {
    const iso = normalizeDate(value)
    if (!iso) return
    const display = isoToDisplayDate(iso)
    console.log(`[AY]   date: raw=${JSON.stringify(value)} → iso="${iso}" → display="${display}"`)
    scrollIntoView(el)
    await focusSetBlur(el, display)
    await sleep(150)
  }

  function fillTextArea (el, value) {
    reactSetTextArea(el, String(value))
  }

  function fillCheckbox (el, desiredBool) {
    if (el.checked === desiredBool) return
    const handled = callReactProp(el, 'onClick',  { target: { checked: desiredBool } }) ||
                    callReactProp(el, 'onChange', { target: { checked: desiredBool } })
    if (!handled) el.click()
  }

  function fillSelectStatic (el, value) {
    const props = getSelectFiberProps(el)
    if (!props) throw new Error(`Select fiber not found for value "${value}"`)
    const norm  = String(value).trim().toLowerCase()
    const match = props.options.find(o =>
      String(o.label ?? o.value ?? '').trim().toLowerCase() === norm
    )
    if (!match) throw new Error(`Option "${value}" not found in list`)
    props.onChange([match.value ?? match.label])
  }

  function waitForOptions (timeoutMs = 5000) {
    return new Promise(resolve => {
      function check () {
        const lb = document.querySelector('[role="listbox"]')
        if (!lb) return null
        const opts = lb.querySelectorAll('[role="option"]')
        return opts.length > 0 ? opts : null
      }
      const existing = check()
      if (existing) { resolve(existing); return }
      const observer = new MutationObserver(() => {
        const result = check()
        if (result) { observer.disconnect(); resolve(result) }
      })
      observer.observe(document.body, { childList: true, subtree: true })
      setTimeout(() => { observer.disconnect(); resolve(null) }, timeoutMs)
    })
  }

  // ─── Human-in-the-loop fill ───────────────────────────────────────────────
  // Used for fields that require the user to pick from a database-linked
  // dropdown (Address, Listing Company, Listing Contact).
  // Copies the value to the clipboard, shows instructions, focuses the
  // field, then polls until the user makes a selection.

  async function fillWithHumanHelp (el, fieldLabel, valueHint, timeoutMs = 300000) {
    const inputEl = el.tagName === 'INPUT' ? el : el.querySelector('input')

    try {
      await navigator.clipboard.writeText(valueHint)
      alert(
        `[ACTION REQUIRED] — ${fieldLabel}\n\n` +
        `"${valueHint}" has been copied to the clipboard.\n\n` +
        `Click OK, paste into the field, and select the correct option from the list.\n` +
        `The script will continue automatically after you make a selection.`
      )
    } catch (_) {
      // Clipboard blocked — use prompt so the user can copy the value manually.
      prompt(
        `[ACTION REQUIRED] — ${fieldLabel}\n\nCopy the value below and paste it into the field:`,
        valueHint
      )
    }

    // Re-focus the field after the dialog closes.
    if (inputEl) {
      inputEl.focus()
      scrollIntoView(inputEl)
    }

    // Wait for the user to select from the dropdown, then ask them to confirm.
    //
    // Why confirm()? Automatic detection is fragile: the field may be
    // pre-filled from a previous session, or the listbox may open/close
    // faster than the polling interval. A confirm() dialog is explicit —
    // the script cannot continue until the user clicks OK, and they can
    // see exactly what value was committed before approving.
    //
    // Flow:
    //   1. Poll every 500ms for: value non-empty + listbox was opened + listbox now closed
    //   2. Show confirm() with the selected value
    //   3. OK → continue  |  Cancel → re-focus and wait for another selection

    const deadline = Date.now() + timeoutMs
    let listboxEverOpened = false

    while (Date.now() < deadline) {
      await sleep(500)
      const val         = inputEl ? inputEl.value.trim() : ''
      const listboxOpen = !!document.querySelector('[role="listbox"]')

      if (val && listboxOpen) listboxEverOpened = true

      if (val && listboxEverOpened && !listboxOpen) {
        const ok = confirm(
          `[${fieldLabel}] — Confirm selection\n\n` +
          `"${val}"\n\n` +
          `OK = continue filling\n` +
          `Cancel = pick again`
        )
        if (ok) {
          console.log(`[AY]   "${fieldLabel}" confirmed: "${val}"`)
          return
        }
        // User wants to retry — reset and re-focus so they can pick again.
        listboxEverOpened = false
        if (inputEl) inputEl.focus()
      }
    }
    throw new Error(`Timeout: "${fieldLabel}" was not filled within ${timeoutMs / 1000}s`)
  }

  async function fillComboboxAsync (el, searchText, opts = {}) {
    const { commitDelay = 600, attempts = 3, listboxTimeout = 5000 } = opts
    const inputEl = el.tagName === 'INPUT' ? el : el.querySelector('input')

    for (let attempt = 0; attempt < attempts; attempt++) {
      // Re-read the fiber each attempt: right after modal open it may not be
      // mounted yet, so we wait up to 1s before giving up on this attempt.
      let sp = getSelectFiberProps(el)
      if (!sp?.onSearch) {
        for (let w = 0; w < 5; w++) {
          await sleep(200)
          sp = getSelectFiberProps(el)
          if (sp?.onSearch) break
        }
      }
      if (!sp?.onSearch) {
        console.warn(`[AY] Attempt ${attempt + 1}/${attempts}: fiber not ready for "${searchText}"`)
        await sleep(300)
        continue
      }

      scrollIntoView(el)
      if (inputEl) {
        inputEl.focus()
        await sleep(80)
        reactSet(inputEl, searchText.trim())
      }
      sp.onSearch(searchText.trim())
      const options = await waitForOptions(listboxTimeout)
      if (!options || options.length === 0) {
        // Address not found in the building database. Don't fall back to
        // free text — the backend requires a valid address_id foreign key.
        throw new Error(`Address not found in building DB: "${searchText}"`)
      }
      const firstOption = options[0]
      console.log(`[AY]   picking first option: "${firstOption.textContent.trim()}"`)
      firstOption.click()
      await sleep(commitDelay)
      if (inputEl) inputEl.blur()
      await sleep(100)
      if (inputEl && inputEl.value.trim()) return
      console.warn(`[AY] Attempt ${attempt + 1}/${attempts}: click didn't commit for "${searchText}"`)
    }
    throw new Error(`Failed to select "${searchText}" after ${attempts} attempts`)
  }

  // ─── 8. Field map ─────────────────────────────────────────────────────────

  const FIELD_MAP = {
    'Address':             { id: 'AddressSelect--0',             type: 'human'             },
    'Suite #':             { id: 'SuiteNumberInput--0',          type: 'text'              },
    'Suite Size':          { id: 'SuiteSizeInput--0',            type: 'number'            },
    'Listing Type':        { id: 'ListingTypeSelect--0',         type: 'select_static'     },
    'Floor':               { id: 'FloorNumInput--0',             type: 'number'            },
    'Warehouse Area':      { id: 'WarehouseAreaInput--0',        type: 'number'            },
    'Office Area':         { id: 'OfficeAreaInput--0',           type: 'number'            },
    'Mezzanine Area':      { id: 'MezzanineAreaInput--0',        type: 'number'            },
    'Yard Area':           { id: 'YardAreaInput--0',             type: 'number'            },
    'Net Rent':            { id: 'NetRentInput--0',              type: 'number'            },
    'Override Gross Rent': { id: 'OverrideGrossRentCheckbox--0', type: 'checkbox'          },
    'Gross Rent':          { id: 'GrossRentInput--0',            type: 'number_if_override'},
    'Rate Type':           { id: 'RateTypeSelect--0',            type: 'select_static'     },
    'Sales Price':         { id: 'SalesPriceInput--0',           type: 'number'            },
    'TI Allowance':        { id: 'TiAllowanceInput--0',          type: 'number'            },
    'Free Rent':           { id: 'FreeRentInput--0',             type: 'number'            },
    'Date Added':          { id: 'DateAddedInput--0',            type: 'date'              },
    'Date Confirmed':      { id: 'DateConfirmedInput--0',        type: 'date'              },
    'Possession Date':     { id: 'PossessionDateInput--0',       type: 'date'              },
    'Availability':        { id: 'AvailabilityInput--0',         type: 'select_static'     },
    'Leased/Sold Date':    { id: 'LeasedSoldDateInput--0',       type: 'date'              },
    'Sublease Expiry':     { id: 'SubleaseExpiryInput--0',       type: 'date'              },
    'Available':           { id: 'AvailableCheckbox--0',         type: 'checkbox'          },
    'Strata/Condo':        { id: 'StrataCheckbox--0',            type: 'checkbox'          },
    'Model Suite':         { id: 'ModalSuiteCheckbox--0',        type: 'checkbox'          },
    'Full Floor':          { id: 'FullFloorCheckbox--0',         type: 'checkbox'          },
    'Improved':            { id: 'ImprovedCheckbox--0',          type: 'checkbox'          },
    'Vacant':              { id: 'VacantCheckbox--0',            type: 'checkbox'          },
    'Under Contract':      { id: 'UnderContractCheckbox--0',     type: 'checkbox'          },
    'Include In Stats':    { id: 'IncludeInStatsCheckbox--0',    type: 'checkbox'          },
    'Space Use Type':      { id: 'SpaceUseTypeSelect--0',        type: 'select_static'     },
    'Ceiling Height':      { id: 'MinHeightInput2--0',           type: 'number'            },
    'Dock Doors':          { id: 'DockInput--0',                 type: 'number'            },
    'Power':               { id: 'PowerInput--0',                type: 'number'            },
    'Grade':               { id: 'GradeInput--0',                type: 'number'            },
    'Divisible':           { id: 'DivisibleInput--0',            type: 'number'            },
    'Contiguous Area':     { id: 'ContigAreaInput--0',           type: 'number'            },
    'Gross Area':          { id: 'GrossAreaInput--0',            type: 'number'            },
    'Use Description':     { id: 'UseDescriptionTypeSelect--0',  type: 'select_static'     },
    'Sort Order':          { id: 'SortOrderInput--0',            type: 'number'            },
    'Listing Company':     { id: 'ListingCompanyInput--0',       type: 'human'             },
    'Listing Contact 1':   { id: 'ListingContact1Input--0',      type: 'human'             },
    'Listing Contact 2':   { id: 'ListingContact2Input--0',      type: 'human'             },
    'Notes':               { id: 'SuiteCommentsInput--0',        type: 'textarea'          },
    'Suite Internal Notes':{ id: 'SuiteInternalNotesInput--0',   type: 'textarea'          },
  }

  function orderedFieldNamesForSuite (suite) {
    const first  = ['Address']
    const late   = ['Listing Company', 'Listing Contact 1', 'Listing Contact 2']
    const middle = Object.keys(suite).filter(
      k => !first.includes(k) && !late.includes(k) && FIELD_MAP[k]
    )
    const tail = late.filter(k => k in suite)
    return [...first.filter(k => k in suite), ...middle, ...tail]
  }

  function toBool (value) {
    if (typeof value === 'boolean') return value
    if (typeof value === 'number')  return value !== 0
    const s = String(value).trim().toLowerCase()
    return s === 'true' || s === 'yes' || s === '1'
  }

  // ─── 9. Fill a single field ───────────────────────────────────────────────

  async function fillField (fieldName, value, suite) {
    const fieldDef = FIELD_MAP[fieldName]
    if (!fieldDef) return

    if (value === null || value === undefined || value === '') {
      console.log(`[AY]   skip "${fieldName}" (empty)`)
      return
    }

    let el = document.getElementById(fieldDef.id)

    // Human-help fields (Listing Company/Contact) only appear in the DOM
    // after Address is selected. Wait up to 3s for them to become visible.
    if (!el && fieldDef.type === 'human') {
      for (let w = 0; w < 15; w++) {
        await sleep(200)
        el = document.getElementById(fieldDef.id)
        if (el) break
      }
    }

    if (!el) {
      console.warn(`[AY]   element not found: "${fieldName}" (id="${fieldDef.id}")`)
      return
    }

    console.log(`[AY]   fill "${fieldName}" →`, value)

    switch (fieldDef.type) {
      case 'text':              await fillText(el, value);          break
      case 'number':            await fillNumber(el, value);        break
      case 'date':              await fillDate(el, value);          break
      case 'textarea':          fillTextArea(el, value);            break
      case 'checkbox':          fillCheckbox(el, toBool(value));    break
      case 'select_static':     fillSelectStatic(el, value);        break
      case 'number_if_override':
        if (toBool(suite['Override Gross Rent'])) {
          await fillNumber(el, value)
        } else {
          console.log(`[AY]   skip "Gross Rent" (Override Gross Rent not checked)`)
        }
        break
      case 'combobox_async': {
        const isContact = fieldName.startsWith('Listing Contact')
        await fillComboboxAsync(el, String(value), { listboxTimeout: isContact ? 5000 : 4000 })
        break
      }
      case 'human':
        await fillWithHumanHelp(el, fieldName, String(value))
        break
      default:
        console.warn(`[AY]   unknown type "${fieldDef.type}" for field "${fieldName}"`)
    }
  }

  // ─── 10. Modal management ─────────────────────────────────────────────────

  async function openNewSuiteModal () {
    const addBtn = [...document.querySelectorAll('button')].find(b => {
      const t = b.textContent.trim()
      return /^(\+|add|new|add row|add suite|add listing)$/i.test(t)
    })
    if (!addBtn) throw new Error('Cannot find "Add" button — is the Retool table visible?')
    addBtn.click()
    for (let i = 0; i < 15; i++) {
      await sleep(200)
      if (document.getElementById('AddressSelect--0')) return
    }
    throw new Error('Modal did not appear after clicking Add button')
  }

  async function closeModalIfOpen () {
    const closeBtn = [...document.querySelectorAll('button')].find(b =>
      /^(close|cancel|×|x)$/i.test(b.textContent.trim())
    )
    if (closeBtn) { closeBtn.click(); await sleep(500) }
  }

  // ─── 11. Form submission ──────────────────────────────────────────────────

  async function submitForm () {
    const btn = [...document.querySelectorAll('button')].find(b =>
      /confirm|save/i.test(b.textContent.trim())
    )
    if (!btn) throw new Error('Cannot find Confirm/Save button')
    btn.click()
    await sleep(1200)

    // Session expiry: a login form appeared after submit.
    if (document.querySelector('input[type="password"]')) {
      throw new Error('Session expired — login page detected after submit. Log in and resume the import.')
    }

    const errorEls = document.querySelectorAll(
      '[class*="error"]:not([style*="display: none"]), [class*="Error"]:not([style*="display: none"])'
    )
    const errors = [...errorEls]
      .map(e => e.textContent.trim())
      .filter(t => t.length > 0 && t.length < 200)
    if (errors.length) throw new Error(`Validation errors: ${errors.join('; ')}`)
  }

  // ─── 12. Review panel ────────────────────────────────────────────────────
  // Shows a non-blocking floating panel so the user can scroll and verify
  // the filled form before it is saved. Returns 'confirm' or 'skip'.

  function waitForReviewConfirmation (suiteLabel) {
    return new Promise((resolve) => {
      const panel = document.createElement('div')
      panel.id = '__ay_review_panel__'
      panel.style.cssText = [
        'position:fixed', 'bottom:24px', 'left:50%', 'transform:translateX(-50%)',
        'z-index:2147483647', 'background:#1e1e2e', 'color:#cdd6f4',
        'font:13px/1.5 monospace', 'padding:14px 20px', 'border-radius:10px',
        'box-shadow:0 4px 28px rgba(0,0,0,.65)', 'display:flex',
        'align-items:center', 'gap:14px', 'pointer-events:all',
        'border:1px solid #313244',
      ].join(';')

      panel.innerHTML = `
        <span style="flex:1">
          <span style="color:#89b4fa;font-weight:700">Review:</span>
          ${suiteLabel} — scroll to check all fields
        </span>
        <button id="__ay_review_confirm__" style="background:#a6e3a1;color:#1e1e2e;border:none;border-radius:6px;padding:8px 18px;font:700 12px monospace;cursor:pointer">
          ✓ Confirm &amp; Save
        </button>
        <button id="__ay_review_skip__" style="background:#f38ba8;color:#1e1e2e;border:none;border-radius:6px;padding:8px 14px;font:700 12px monospace;cursor:pointer">
          ✗ Skip
        </button>
      `

      document.body.appendChild(panel)

      document.getElementById('__ay_review_confirm__').addEventListener('click', () => {
        panel.remove()
        resolve('confirm')
      })
      document.getElementById('__ay_review_skip__').addEventListener('click', () => {
        panel.remove()
        resolve('skip')
      })
    })
  }

  // ─── 13. Suite-level fill ─────────────────────────────────────────────────

  async function fillSuite (suite) {
    await openNewSuiteModal()
    const fieldNames = orderedFieldNamesForSuite(suite)
    for (const fieldName of fieldNames) {
      try {
        await fillField(fieldName, suite[fieldName], suite)
        await sleep(150)
      } catch (err) {
        console.warn(`[AY]   field error "${fieldName}":`, err.message)
      }
    }

    if (SETTINGS.dryRun) {
      console.log('[AY]   [DRY RUN] Fields filled — closing modal without saving')
      await closeModalIfOpen()
      return
    }

    if (!SETTINGS.submitAfterFill) return

    if (SETTINGS.reviewBeforeSave) {
      const label = suite['Address'] || suite['Suite #'] || 'suite'
      console.log(`[AY]   Waiting for user review of "${label}"…`)
      const decision = await waitForReviewConfirmation(label)
      if (decision === 'skip') {
        console.log(`[AY]   Skipped (user chose not to save "${label}")`)
        await closeModalIfOpen()
        return
      }
      console.log(`[AY]   User confirmed — saving "${label}"`)
    }

    await submitForm()
  }

  // ─── 14. Main import loop ─────────────────────────────────────────────────

  async function importSuites (suites, startIndex = 0) {
    const total   = suites.length
    const results = []

    overlayCreate(total)
    console.log(`[AY] Starting import: ${total} suite(s), from index ${startIndex}`)
    if (SETTINGS.dryRun) console.log('[AY] ⚠ DRY RUN active — forms will be filled but NOT saved')

    // Restore results for already-completed suites when resuming.
    for (let i = 0; i < startIndex; i++) {
      results.push({ index: i, label: suites[i]['Address'] || suites[i]['Suite #'] || `row ${i+1}`, status: 'skipped (resumed)' })
    }

    for (let i = startIndex; i < total; i++) {
      const suite = suites[i]
      const label = suite['Address'] || suite['Suite #'] || `row ${i + 1}`
      console.log(`[AY] Suite ${i + 1}/${total}: ${label}`)
      overlayUpdate({ done: i, total, label, status: null })

      let status = 'success'
      let errorMsg = null

      try {
        await fillSuite(suite)
        console.log(`[AY] ✓ Suite ${i + 1}/${total}: ${label}`)
      } catch (err) {
        status   = 'error'
        errorMsg = err.message
        console.error(`[AY] ✗ Suite ${i + 1}/${total} failed: ${err.message}`)
        await closeModalIfOpen()
      }

      results.push({ index: i, label, status, error: errorMsg, ts: new Date().toISOString() })
      overlayUpdate({ done: i + 1, total, label, status })

      // Save checkpoint after every suite.
      sessionSave({ total, lastCompletedIndex: i, results })

      if (status === 'error' && !SETTINGS.skipErrors) {
        console.error('[AY] Stopped — set SETTINGS.skipErrors=true to continue on errors.')
        overlayDestroy(`Stopped on error — suite ${i + 1}`)
        return results
      }

      if (i < total - 1) {
        const next      = suites[i + 1]
        const nextLabel = next['Address'] || next['Suite #'] || `row ${i + 2}`
        const go = confirm(`Suite ${i + 1}/${total} done.\n\nContinue to next?\n→ ${nextLabel}`)
        if (!go) {
          console.log('[AY] Import paused by user.')
          overlayDestroy('Paused by user')
          return results
        }
        await sleep(SETTINGS.delay)
      }
    }

    sessionClear()
    overlayDestroy(`Done — ${total} suite(s)`)
    console.log('[AY] Import complete')
    return results
  }

  // ─── Pre-flight check ────────────────────────────────────────────────────
  // Runs before any file is picked. Aborts early with a clear message if the
  // environment is not ready so we don't waste the user's time.

  function preflight () {
    const issues = []

    // Must be on a Retool app page. Retool apps always use the path pattern
    // /app/<name>/page<n> regardless of the host domain, so we check the path
    // instead of the hostname — this supports custom domains like
    // apps.hyperdrive.avisonyoung.com as well as *.retool.com.
    if (!/^\/app\/[^/]+\/page/.test(location.pathname)) {
      issues.push(
        'Page does not appear to be a Retool app — navigate to the Availability Manager ' +
        '(expected URL path: /app/<name>/page<n>, current: ' + location.pathname + ')'
      )
    }

    // The "Add" button must be visible — confirms the table is loaded.
    const addBtn = [...document.querySelectorAll('button')].find(b =>
      /^(\+|add|new|add row|add suite|add listing)$/i.test(b.textContent.trim())
    )
    if (!addBtn) {
      issues.push('"Add / +" button not found — is the suites table visible?')
    }

    // Must not already be showing a login form.
    if (document.querySelector('input[type="password"]')) {
      issues.push('Session appears to have expired — a password field is visible on the page')
    }

    if (issues.length) {
      const msg = 'PRE-FLIGHT FAILED:\n\n' + issues.map((v, i) => `${i + 1}. ${v}`).join('\n')
      alert(msg)
      throw new Error('Pre-flight failed: ' + issues.join('; '))
    }

    console.log('[AY] Pre-flight OK')
  }

  // ─── Data validation ──────────────────────────────────────────────────────
  // Checks all parsed suites before starting so the user knows about problems
  // upfront rather than discovering them mid-import.

  const REQUIRED_FIELDS = ['Address', 'Suite #']
  const DATE_FIELDS     = ['Date Added', 'Date Confirmed', 'Possession Date', 'Leased/Sold Date', 'Sublease Expiry']

  function validateSuites (suites) {
    const warnings = []

    suites.forEach((suite, i) => {
      const row = i + 1

      // Required fields.
      for (const f of REQUIRED_FIELDS) {
        const v = String(suite[f] ?? '').trim()
        if (!v) warnings.push(`Row ${row}: required field "${f}" is empty`)
      }

      // Date fields — must be parseable.
      for (const f of DATE_FIELDS) {
        const v = suite[f]
        if (v === undefined || v === null || v === '') continue
        const iso = normalizeDate(v)
        if (!iso || iso.length < 8) {
          warnings.push(`Row ${row}: "${f}" = "${v}" is not a valid date`)
        }
      }
    })

    if (warnings.length === 0) {
      console.log('[AY] Validation OK — no issues found')
      return true
    }

    const preview = warnings.slice(0, 10).join('\n')
    const extra   = warnings.length > 10 ? `\n…and ${warnings.length - 10} more issue(s)` : ''
    const go = confirm(
      `VALIDATION FOUND ${warnings.length} ISSUE(S):\n\n${preview}${extra}\n\n` +
      `OK = continue anyway\nCancel = abort`
    )
    return go
  }

  // ─── 15. Entry point ──────────────────────────────────────────────────────

  try {
    console.log('[AY] Loading SheetJS…')
    await loadSheetJS()
    preflight()

    // ── Check for a saved session and offer to resume ──────────────────────
    const saved = sessionLoad()
    if (saved) {
      const { total, lastCompletedIndex, results } = saved
      const done = lastCompletedIndex + 1
      const resume = confirm(
        `Previous session found!\n\n` +
        `${done}/${total} suite(s) completed.\n\n` +
        `OK = resume from suite ${done + 1}\n` +
        `Cancel = start over (you will need to pick the file again)`
      )
      if (resume) {
        console.log(`[AY] Resuming from suite ${done + 1}/${total}…`)
        console.log('[AY] SheetJS ready. Opening file picker…')
        const { name, buffer } = await pickFile()
        const wb     = XLSX.read(new Uint8Array(buffer), { type: 'array', cellDates: false })
        const suites = parseWorkbook(wb)
        if (suites.length !== total) {
          alert(`Warning: file has ${suites.length} rows but the session expected ${total}. Starting over.`)
          sessionClear()
        } else {
          await importSuites(suites, done)
          return
        }
      } else {
        sessionClear()
      }
    }

    console.log('[AY] SheetJS ready. Opening file picker…')
    const { name, buffer } = await pickFile()
    console.log(`[AY] File: "${name}". Parsing…`)

    const wb     = XLSX.read(new Uint8Array(buffer), { type: 'array', cellDates: false })
    const suites = parseWorkbook(wb)

    if (suites.length === 0) {
      console.error('[AY] No rows found in the spreadsheet.')
      return
    }

    console.log(`[AY] ${suites.length} suite(s) found. Starting import…`)
    await importSuites(suites)
  } catch (err) {
    console.error('[AY] Fatal error:', err.message)
  }

})()
