// import ExcelJS from "https://dbm-u.mgamecorp.com/js/exceljs.cdn.mzzang.v1.js";
// import * as ExcelJS from "https://dbm-u.mgamecorp.com/js/exceljs.cdn.mzzang.v1.js";
import "https://dbm-u.mgamecorp.com/js/exceljs.cdn.mzzang.v1.js";
const ExcelJS = window.ExcelJS;

export { _RenderingList_v2, scroll_process_class, _excelCellExport };

// ─── 가상 스크롤 ───────────────────────────────────────────────
class VirtualScroll {
   constructor(bodyEl, listArr, exception_col, evenoddflag, options = {}) {
      this.bodyEl     = bodyEl;
      this.list       = listArr.list;
      this.evenoddFlag = evenoddflag ?? true;
      this.excCol     = exception_col;
      this.spacerWidth = options.spacerWidth ?? null;
      this._spacerWidthSet = false;
      this.estimatedH = options.rowHeight ?? 32;
      this.buffer     = options.buffer ?? 15;
      this.colCnt     = options.colCnt ?? 0;   // ← sticky 컬럼 수
      this.rendered   = new Map();
      this.heights    = new Array(this.list.length).fill(this.estimatedH);
      this.positions  = this._buildPositions();
      this.leftPos    = null;  // ← sticky left 위치 캐시 (최초 렌더 시 계산)

      this._setup();
   }

   _buildPositions() {
      const pos = new Array(this.list.length + 1);
      pos[0] = 0;
      for (let i = 0; i < this.list.length; i++)
         pos[i + 1] = pos[i] + this.heights[i];
      return pos;
   }

   _findStartIdx(scrollTop) {
      let lo = 0, hi = this.list.length - 1;
      while (lo < hi) {
         const mid = (lo + hi) >> 1;
         if (this.positions[mid + 1] <= scrollTop) lo = mid + 1;
         else hi = mid;
      }
      return lo;
   }

   _setup() {

      // 가상스크롤 even/odd CSS (nth-child보다 높은 우선순위로 덮어씀)
      if( this.evenoddFlag ) {
         if (!document.querySelector('style[data-vs-evenodd]')) {
            const style = document.createElement('style');
            style.dataset.vsEvenodd = '';
            style.textContent = `
               .list_body .list_item.vs-row-even { background-color: var(--row-even, #ffffff) !important; }
               .list_body .list_item.vs-row-odd  { background-color: var(--row-odd,  #ebebeb) !important; }
            `;
            document.head.appendChild(style);
         }
      }
      
      this.spacer = document.createElement('div');
      this.spacer.style.height = `${this.positions[this.list.length]}px`;

      // 픽셀값이면 바로 적용, 아니면 렌더 후 측정
      if (this.spacerWidth && /px$/.test(this.spacerWidth)) {
         this.spacer.style.width = this.spacerWidth;
         this._spacerWidthSet = true;
      }
      
      this.bodyEl.style.position = 'relative';
      this.bodyEl.innerHTML = '';
      this.bodyEl.appendChild(this.spacer);

      this._render();
      this._handler = () => this._render();
      window.addEventListener('scroll', this._handler, { passive: true });
   }

   _buildRowEl(rowData, index) {
      const { _style: styles, _class: classes, _licontent: contents, _tooltip: tooltips } = rowData;
      const items = styles.map((style, k) => {
         const hasLink  = /<a[\s>]/i.test(contents[k]);
         const hasInput = /<(input|select)[\s>]/i.test(contents[k]);
         const title = hasInput 
            ? '' 
            : (tooltips?.[k] ?? _extractText(contents[k]).replaceAll("'", ''));
         const cls      = classes?.[k] ? ` class="${classes[k]}"` : '';
         const content_col = (!hasLink && !hasInput && contents[k] &&
            (this.excCol === undefined || k >= this.excCol))
            ? `<span class="inline-tooltip-wrapper" data-tooltip="${title}">${contents[k]}</span>`
            : contents[k];
         return `<li${cls} style="${style}" aria-label="${title}">${content_col}</li>`;
      });

      // const ul = document.createElement('ul');
      // ul.className = 'list_item';
      // ul.innerHTML  = items.join('');
      // ul.style.cssText = `position:absolute;top:${this.positions[index]}px;width:100%;`;
      // return ul;
      const ul = document.createElement('ul');
      ul.className = 'list_item';
      ul.innerHTML = items.join('');
      ul.classList.add(index % 2 === 0 ? 'vs-row-even' : 'vs-row-odd');  // ← 클래스로
      ul.style.cssText = `position:absolute;top:${this.positions[index]}px;width:100%;`;
      // background-color 인라인 제거
      return ul;
   }

   // DOM에 붙은 후 offsetWidth 읽기 (최초 1회)
   _calcLeftPos(row) {
      const cols = [...row.querySelectorAll('li')].slice(0, this.colCnt);
      let acc = 0;
      return cols.map(col => {
         const left = acc;
         acc += col.offsetWidth;
         return left;
      });
   }

   // 단일 행에 sticky 스타일 적용
   _applyStickyRow(ul) {
      [...ul.querySelectorAll('li')].forEach((col, i) => {
         if (i >= this.colCnt) return;
         const isLast = i === this.colCnt - 1;
         col.classList.add('sticky-col');
         col.style.setProperty('position', 'sticky');
         col.style.setProperty('left', `${this.leftPos[i]}px`, 'important');
         // col.style.setProperty('background-color', 'rgba(223, 230, 234, 1)');
         col.style.setProperty('background-color', 'var(--list-head-bg)');
         col.style.setProperty('z-index', '7');
         if (isLast) col.style.setProperty('box-shadow', '7px 0 7px -5px rgba(0,0,0,0.3)');
      });
   }

   // 헤더 sticky 적용 (최초 1회)
   _applyStickyHead() {
      for (let i = 0; i < this.colCnt; i++) {
         const isLast = i === this.colCnt - 1;
         ['.list_head', '.list_head_v2', '.list_head_sub'].forEach(sel => {
            const headLi = document.querySelector(`${sel} li:nth-child(${i + 1})`);
            if (!headLi) return;
            headLi.classList.add('sticky-col');
            headLi.style.setProperty('position', 'sticky');
            headLi.style.setProperty('left', `${this.leftPos[i]}px`, 'important');
            headLi.style.setProperty('z-index', '8');
            if (isLast) headLi.style.setProperty('box-shadow', '7px 0 7px -5px rgba(0,0,0,0.3)');
         });
      }
   }

   _render() {
      const containerTop = this.bodyEl.getBoundingClientRect().top + window.scrollY;
      const relScroll    = Math.max(0, window.scrollY - containerTop);

      const startIdx = Math.max(0, this._findStartIdx(relScroll) - this.buffer);
      let endIdx = startIdx;
      while (endIdx < this.list.length && this.positions[endIdx] < relScroll + window.innerHeight)
         endIdx++;
      endIdx = Math.min(this.list.length, endIdx + this.buffer);

      // 범위 밖 행 제거
      for (const [i, el] of this.rendered) {
         if (i < startIdx || i >= endIdx) {
            el.remove();
            this.rendered.delete(i);
         }
      }

      // 새 행 추가
      const newIndices = [];
      const fragment = document.createDocumentFragment();
      for (let i = startIdx; i < endIdx; i++) {
         if (this.rendered.has(i)) continue;
         const el = this._buildRowEl(this.list[i], i);
         fragment.appendChild(el);
         this.rendered.set(i, el);
         newIndices.push(i);
      }
      this.bodyEl.appendChild(fragment);  // DOM에 붙이기

      // spacer 너비 미설정 시 첫 행에서 측정
      if (!this._spacerWidthSet && newIndices.length > 0) {
         const firstEl = this.rendered.get(newIndices[0]);
         const w = firstEl.scrollWidth;  // 가로 스크롤 포함 전체 너비
         if (w > 0) {
            this.spacer.style.width = `${w}px`;
            this._spacerWidthSet = true;
         }
      }

      if (!newIndices.length) return;

      // ─── sticky 처리 ─────────────────────────────────────────
      if (this.colCnt) {
         if (!this.leftPos) {
            // 최초: leftPos 계산 후 헤더 + 전체 렌더 행에 적용
            this.leftPos = this._calcLeftPos(this.rendered.get(newIndices[0]));
            this._applyStickyHead();
            for (const [, el] of this.rendered) this._applyStickyRow(el);
         } else {
            // 이후: 새로 추가된 행에만 적용
            for (const i of newIndices) this._applyStickyRow(this.rendered.get(i));
         }
      }

      // ─── 높이 측정 및 위치 보정 ──────────────────────────────
      let changed = false;
      for (const i of newIndices) {
         // const h = this.rendered.get(i).offsetHeight;
         const el = this.rendered.get(i);
         const liEls = [...el.querySelectorAll('li')];
         const h = liEls.length > 0
            ? Math.ceil(Math.max(...liEls.map(li => li.getBoundingClientRect().height)))
            : el.offsetHeight;

         if (h > 0 && h !== this.heights[i]) {
            this.heights[i] = h + 1;
            changed = true;
         }
      }
      if (changed) {
         this.positions = this._buildPositions();
         this.spacer.style.height = `${this.positions[this.list.length]}px`;
         for (const [i, el] of this.rendered)
            el.style.top = `${this.positions[i]}px`;
      }
   }

   destroy() {
      window.removeEventListener('scroll', this._handler);
      this.rendered.clear();
   }
}


// ─── 공통 유틸 ───────────────────────────────────────────────
const _decodeEl = document.createElement('div');

function _extractText(html) {
   if (!html) return '';
   _decodeEl.innerHTML = html;
   return _decodeEl.textContent;
}

function deepChangeOption(target, source) {
   for (const key of Object.keys(source)) {
      const val = source[key];
      if (val && typeof val === 'object' && !Array.isArray(val)
          && !(val instanceof Date) && !(val instanceof RegExp)) {
         if (!target[key]) target[key] = {};
         deepChangeOption(target[key], val);
      } else {
         target[key] = val;
      }
   }
   return target;
}

function colorNameToARGB(color) {
   if (!color) return 'FF000000';
   if (color.startsWith('#')) return 'FF' + color.replace('#', '').toUpperCase();
   const map = {
      red: 'FFFF0000', blue: 'FF0000FF', green: 'FF00FF00',
      black: 'FF000000', white: 'FFFFFFFF', yellow: 'FFFFFF00',
   };
   return map[color.toLowerCase()] ?? 'FF000000';
}

// 테두리 강제 적용 헬퍼
function applyBorderToRange(ws, startRow, startCol, endRow, endCol, border) {
   for (let r = startRow; r <= endRow; r++)
      for (let c = startCol; c <= endCol; c++)
         ws.getCell(r, c).border = border;
}

// ✅ 버그 수정: color 캡처 그룹 복원 (match[1]=color, match[2]=text)
function parseHtmlToRichText(html, ftsize) {
   const regex = /<font\s+color=["']?(#[0-9a-fA-F]{6}|[a-zA-Z]+)["']?>(.*?)<\/font>/gis;
   const defaultFont = { size: ftsize, color: { argb: 'FF000000' } };
   const richText = [];
   let lastIndex = 0, match;

   while ((match = regex.exec(html)) !== null) {
      if (match.index > lastIndex)
         richText.push({ text: html.substring(lastIndex, match.index), font: { ...defaultFont } });
      richText.push({
         text: match[2],
         font: { ...defaultFont, color: { argb: colorNameToARGB(match[1]) } },
      });
      lastIndex = regex.lastIndex;
   }
   if (lastIndex < html.length)
      richText.push({
         text: html.substring(lastIndex).replace(/<\/?font[^>]*>/gi, ''),
         font: { ...defaultFont },
      });
   return richText;
}

function _stripHtml(str) {
   return str.replace(/<[^>]*>/g, '').trim();
}

// ─── 클립보드 ────────────────────────────────────────────────

async function _copyText(text) {
   try {
      await navigator.clipboard.writeText(text);
      const el = document.querySelector('.copy-feedback');
      if (el) { el.classList.remove('failtxt'); el.textContent = 'copied!'; }
   } catch {
      const el = document.querySelector('.copy-feedback');
      if (el) { el.classList.add('failtxt'); el.textContent = 'fail copied..'; }
   }
}

function _showCopyFeedback(copyEl, listDoc) {
   if (!copyEl.textContent.trim()) return;
   const rect = copyEl.getBoundingClientRect();
   const y = rect.top + window.scrollY - (copyEl.clientHeight + 3);
   const x = rect.left + window.scrollX + copyEl.clientWidth / 2;
   document.querySelector('.copy-feedback')?.remove();
   document.querySelector(`.list_body${listDoc}`)
      .insertAdjacentHTML('beforebegin',
         `<div class="copy-feedback" style="top:${y}px;left:${x}px">copied!</div>`);
   _copyText(copyEl.textContent.trim());
}


// ─── 리스트 렌더링 헬퍼 ──────────────────────────────────────

function _setupListCss(listDoc, styleRef) {
   let styleEl = document.querySelector(`style[data-list="${listDoc}"]`);
   if (!styleEl) {
      styleEl = document.createElement('style');
      styleEl.dataset.list = listDoc;
      document.head.appendChild(styleEl);
   }
   let css = '';
   if (styleRef.rowHoverBlur) css += `
      .list_body${listDoc}:hover ul { filter: blur(2px); }
      .list_body${listDoc} ul:hover { filter: none; }
   `;
   if (styleRef.evenoddFlag ?? true) css += `
      .list_item:nth-child(even) { background-color: var(--row-even); }
      .list_item:nth-child(odd)  { background-color: var(--row-odd); }
   `;

   styleEl.textContent = css;
}

function _normalizeWidth(w) {
   const s = String(w);
   return /[%px]|max-content/.test(s) ? s : s + 'px';
}

function _buildHeadHtml({ _sbj: cols, _class: classes, _tooltip: tooltips, _style: styles }) {
   const items = cols.map((col, i) => {
      const cls = classes?.[i] ? `class="${classes[i]}"` : '';
      const hasLink = /<a[\s>]/i.test(col);
      const content_col = (!hasLink && tooltips?.[i])
         ? `<span class="inline-tooltip-wrapper" data-tooltip="${tooltips[i]}">${col}</span>`
         : col;
      return `<li ${cls} style="${styles[i]}" aria-label="${_stripHtml(col)}">${content_col}</li>`;
   });
   return `<ul class="list_head_item">${items.join('')}</ul>`;
}

function _buildSubHeadHtml({ _sbj: cols, _class: classes, _tooltip: tooltips, _style: styles }, width, topsize, fsize) {
   const items = cols.map((col, i) => {
      const cls = classes?.[i] ? `class="${classes[i]}"` : '';
      const hasLink = /<a[\s>]/i.test(col);
      const content_col = (!hasLink && tooltips?.[i])
         ? `<span class="inline-tooltip-wrapper" data-tooltip="${tooltips[i]}">${col}</span>`
         : col;
      return `<li ${cls} style="${styles[i]}" aria-label="${_stripHtml(col)}">${content_col}</li>`;
   });
   return `<div class="list_head_sub" style="width:${width};font-size:${fsize}px;top:${topsize}px">
      <ul class="list_head_item">${items.join('')}</ul>
   </div>`;
}

function _buildBodyHtml(listArr, exception_col) {
   if (listArr.rstCode === -1)
      return `<ul class="list_item"><li class="nodata">검색된 데이터가 없습니다. (1)</li></ul>`;
   if (listArr.rstCode !== 1 || !listArr.list || listArr.list === '')
      return `<ul class="list_item"><li class="nodata">검색된 데이터가 없습니다. (2)</li></ul>`;

   return listArr.list.map(({ _style: styles, _class: classes, _licontent: contents }) => {
      const items = styles.map((style, k) => {
         const hasLink = /<a[\s>]/i.test(contents[k]);
         const hasInput = /<(input|select)[\s>]/i.test(contents[k]);
         const title = hasInput ? '' : _extractText(contents[k]).replaceAll("'", '');

         const cls = classes?.[k] ? `class="${classes[k]}"` : '';
         const content_col = (!hasLink && !hasInput && contents[k] && (exception_col === undefined || k >= exception_col) )
            ? `<span class="inline-tooltip-wrapper" data-tooltip="${title}">${contents[k]}</span>`
            : contents[k];
         return `<li ${cls} style="${style}" aria-label="${title}">${content_col}</li>`;
      });
      return `<ul class="list_item">${items.join('')}</ul>`;
   }).join('');
}

function _attachHoverHighlight(listDoc, controller) {
   const listBody = document.querySelector(`.list_body${listDoc}`);

   listBody.addEventListener('mouseover', e => {
      listBody.querySelectorAll('.mhover').forEach(el => el.classList.remove('mhover'));
      document.querySelectorAll('.list_head .mhover, .list_head_v2 .mhover, .list_head_sub .mhover')
         .forEach(el => el.classList.remove('mhover'));

      const liHover = e.target.closest(`.list_body${listDoc} li`);
      if (!liHover) return;

      const ulHover = e.target.closest('ul.list_item');
      // rowIdx 불필요 — ulHover에서 바로 sticky-col 탐색
      ulHover?.querySelectorAll('li.sticky-col').forEach(el => el.classList.add('mhover'));

      if (!liHover.classList.contains('sticky-col')) {
         const colIdx = [...liHover.parentElement.querySelectorAll('li')].indexOf(liHover);
         const headhover_ele = document.querySelector('.list_head_sub') ?? document.querySelector('.list_head');
         headhover_ele?.querySelector(`li:nth-child(${colIdx + 1})`)?.classList.add('mhover');
      }
   }, { signal: controller.signal });

   listBody.addEventListener('mouseleave', () => {
      document.querySelectorAll('.mhover').forEach(el => el.classList.remove('mhover'));
   }, { signal: controller.signal });
}


// ─── 툴팁 ────────────────────────────────────────────────────
function _attachHoverTooltip(tooltipFlag) {

   // v3 이벤트 다수 발생 문제로 개선
   // 중복 방지
   if (document.body._tooltipAttached) {
      // title → data-tooltip 변환만 수행
      document.querySelectorAll('[title]').forEach(zone => {
         zone.setAttribute('data-tooltip', zone.getAttribute('title'));
         zone.removeAttribute('title');
      });
      return;
   }
   document.body._tooltipAttached = true;

   const box = document.createElement('span');
   box.className = 'tooltip-box';
   document.body.appendChild(box);
   const OFFSET = 7;

   // title → data-tooltip 사전 변환 (위임 시 data-tooltip만 읽으면 되도록)
   const selector = tooltipFlag.bodyTooltip === true
      ? '[data-tooltip], [title]'
      : '.list_head_item [data-tooltip], .list_head_item [title]';

   document.querySelectorAll(selector).forEach(zone => {
      const title = zone.getAttribute('title');
      if (title) {
         zone.setAttribute('data-tooltip', title);
         zone.removeAttribute('title');
      }
   });

   // 리스너 3개만 (body에 위임)
   document.body.addEventListener('mouseover', (e) => {
      const zone = e.target.closest('[data-tooltip]');
      if (!zone) return;
      if (!tooltipFlag.bodyTooltip && !zone.closest('.list_head_item')) return;
      box.textContent = zone.getAttribute('data-tooltip');
      box.style.opacity = '1';
   });

   document.body.addEventListener('mouseout', (e) => {
      const zone = e.target.closest('[data-tooltip]');
      if (zone && !zone.contains(e.relatedTarget)) box.style.opacity = '0';
   });

   document.body.addEventListener('mousemove', (e) => {
      if (box.style.opacity !== '1') return;
      const boxW = box.offsetWidth;
      const boxH = box.offsetHeight;
      const x = e.clientX + OFFSET + boxW > window.innerWidth - 150
         ? e.clientX - OFFSET - boxW : e.clientX + OFFSET;
      const y = e.clientY + OFFSET + boxH > window.innerHeight - 150
         ? e.clientY - OFFSET - boxH : e.clientY + OFFSET;
      box.style.left = x + 'px';
      box.style.top  = y + 'px';
   });
}

function _attachCopyCol(listDoc, controller, copyColIndices) {
   document.querySelector(`.list_body${listDoc}`)?.addEventListener('click', e => {
      copyColIndices.forEach(idx => {
         const col = e.target.closest(`ul.list_item li:nth-child(${idx})`);
         if (!col) return;
         e.preventDefault();
         _showCopyFeedback(col, listDoc);
      });
   }, { signal: controller.signal });
}


// ─── _RenderingList_v2 ───────────────────────────────────────

const _RenderingList_v2 = (list_head, list_head_sub, list_arr, styleRef, list_doc = '') => {

   // styleRef : width, fsize, rowHoverBlur, evenoddFlag, searchResize, headhover, copycol, horizonscrollcolcnt, bodyTooltip

   window[`_listController${list_doc}`]?.abort();
   const controller = new AbortController();
   window[`_listController${list_doc}`] = controller;

   const width = _normalizeWidth(styleRef?.width ?? '100%');
   const fsize = styleRef?.fsize ?? 12;

   _setupListCss(list_doc, styleRef);

   if (document.querySelector('.dbm2025_search_list') && styleRef.searchResize)
      document.querySelector('.dbm2025_search_list').style.width = width;
   document.querySelector('.list-top') && (document.querySelector('.list-top').style.width = width);
   document.querySelectorAll('.list_item').forEach(e => (e.style.width = width));

   const headEl = document.querySelector(`.list_head${list_doc}`);
   headEl.style.width    = width;
   headEl.style.fontSize = fsize + 'px';
   headEl.innerHTML      = _buildHeadHtml(list_head);

   if (list_head_sub?._sbj?.length) {
      headEl.classList.add('list_head_v2');
      headEl.classList.remove('list_head');
      // console.log(document.querySelector(`.list_head_v2`).style);
      const currentHeadEl_height = document.querySelector(`.list_head_v2`).offsetHeight;
      document.querySelector('.list_head_sub')?.remove();
      headEl.insertAdjacentHTML('afterend', _buildSubHeadHtml(list_head_sub, width, currentHeadEl_height, fsize));
   }

   const bodyEl = document.querySelector(`.list_body${list_doc}`);
   bodyEl.style.width    = width;
   bodyEl.style.fontSize = fsize + 'px';
   window[`_vs${list_doc}`]?.destroy();

   if (list_arr.rstCode === 1 && list_arr.list?.length) {
      window[`_vs${list_doc}`] = new VirtualScroll(
         bodyEl, list_arr, styleRef.horizonscrollcolcnt, styleRef.evenoddFlag,
         {
            rowHeight: styleRef.rowHeight ?? 35,
            colCnt:    styleRef.horizonscrollcolcnt ?? 0,
         }
      );
   } else {
      bodyEl.innerHTML = _buildBodyHtml(list_arr, styleRef.horizonscrollcolcnt);
   }

   requestAnimationFrame(() => {
      if (styleRef.headhover ?? false) _attachHoverHighlight(list_doc, controller);
      if (styleRef.copycol)            _attachCopyCol(list_doc, controller, styleRef.copycol);
      _attachHoverTooltip(styleRef);

      const hasVS = !!(window[`_vs${list_doc}`]);
      scroll_process_class(styleRef.horizonscrollcolcnt, hasVS); // ← hasVS이면 body sticky 스킵
   });
};

// ─── 스크롤 / 스티키 ─────────────────────────────────────────

const scroll_process_class = (horizonColCnt, skipStickyPos = false) => {
   const headers = document.querySelectorAll('.list_head, .list_head_v2, .list_head_sub');
   if (!headers.length) return;

   headers.forEach(header => {
      if (header._stickyHandler) window.removeEventListener('scroll', header._stickyHandler);

      header.classList.remove('is-sticky');  // ← 추가: 측정 전 초기화
      const initialOffset = header.getBoundingClientRect().top + window.pageYOffset;

      const handler = () => header.classList.toggle('is-sticky', window.pageYOffset > initialOffset);
      header._stickyHandler = handler;
      window.addEventListener('scroll', handler, { passive: true });
      handler();  // 현재 스크롤 위치 기준으로 즉시 재적용
   });

   // // VirtualScroll 사용 시 body 행 sticky는 내부에서 처리하므로 스킵
   if (!skipStickyPos) update_sticky_positions(horizonColCnt);
};

const update_sticky_positions = (colCnt) => {
   // v2 페이지 속도 문제로 개선
   if (!colCnt) return;
   const rows = document.querySelectorAll('.list_item');
   if (!rows.length) return;

   // 1단계: 쓰기 전에 너비 일괄 읽기 (reflow 최소화)
   const widths = [...rows[0].querySelectorAll('li')]
      .slice(0, colCnt)
      .map(col => col.offsetWidth);

   const leftPositions = [];
   let accLeft = 0;
   for (let i = 0; i < colCnt; i++) {
      leftPositions.push(accLeft);
      accLeft += widths[i];
   }

   // 2단계: 읽기 없이 쓰기만
   rows.forEach(row => {
      [...row.querySelectorAll('li')].forEach((col, index) => {
         if (index >= colCnt) return;
         const isLast = index === colCnt - 1;
         col.classList.add('sticky-col');
         col.style.setProperty('position', 'sticky');
         col.style.setProperty('left', `${leftPositions[index]}px`, 'important');
         col.style.setProperty('background-color', 'rgba(223, 230, 234, 1)');
         col.style.setProperty('z-index', '7');
         if (isLast) col.style.setProperty('box-shadow', '7px 0 7px -5px rgba(0,0,0,0.3)');
      });
   });

   // 헤더
   ['.list_head', '.list_head_v2', '.list_head_sub'].forEach(sel => {
      for (let i = 0; i < colCnt; i++) {
         const headLi = document.querySelector(`${sel} li:nth-child(${i + 1})`);
         if (!headLi) continue;
         const isLast = i === colCnt - 1;
         headLi.classList.add('sticky-col');
         headLi.style.setProperty('position', 'sticky');
         headLi.style.setProperty('left', `${leftPositions[i]}px`, 'important');
         headLi.style.setProperty('z-index', '8');
         if (isLast) headLi.style.setProperty('box-shadow', '7px 0 7px -5px rgba(0,0,0,0.3)');
      }
   });
};


// ─── 엑셀 내보내기 ───────────────────────────────────────────

// border type
// | 스타일 이름         | 설명
// | ------------------ | -------------------------------------
// | `thin`             | 얇은 선 (기본적으로 가장 많이 사용)
// | `dotted`           | 점선 (········)
// | `dashDot`          | 점-선 혼합 (— · — · —)
// | `dashDotDot`       | 점-점-선 혼합 (— · · — · ·)
// | `dashed`           | 대시 선 (— — —)       
// | `double`           | 이중선 (=)
// | `hair`             | 매우 얇은 선 (thin보다 더 얇음)
// | `medium`           | 중간 굵기 선
// | `mediumDashed`     | 중간 굵기의 대시 선
// | `mediumDashDot`    | 중간 굵기의 점-선
// | `mediumDashDotDot` | 중간 굵기의 점-점-선
// | `slantDashDot`     | 사선 대시 점선 (지원 여부는 프로그램에 따라 다름)
// | `none`             | 테두리 없음
// top, bottom 중 하나라도 thin인 경우 thin으로 나옴. 한쪽을 none으로 해야 원하는 스타일이 적용됨. (left, right도 동일)

// font type
// | 스타일 이름      | 값 예시                | 설명
// | --------------- | --------------------- |------------------------------
// | `name`          | 'Calibri',            | 폰트 이름 (기본: Calibri)
// | `family`        | 2,                    | 폰트 패밀리 (1=Roman, 2=Swiss, 3=Modern, 4=Script, 5=Decorative)
// | `size`          | 12,                   | 폰트 크기
// | `bold`          | true,                 | 굵게
// | `italic`        | true,                 | 기울임꼴
// | `underline`     | true,                 | 밑줄 (true, false, 'double', 'singleAccounting', 'doubleAccounting')
// | `strike`        | true,                 | 취소선
// | `color`         | { argb: 'FF0000FF' }  | 색상 (ARGB: Alpha-Red-Green-Blue)
// | `outline`       | true,                 | 아웃라인 폰트
// | `shadow`        | true,                 | 그림자
// | `vertAlign`     | 'superscript'         | 위첨자 ('superscript') 또는 아래첨자 ('subscript')

// styleOption 예제 값
// {
//    mainStyle: {
//       font: { bold: true, size: 10, color: { argb: '000000' } },
//       fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'dfe6ea' } },
//       border: { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} },
//       alignment: { horizontal: 'center', vertical: 'middle' },
//       rowHeight: 25,
//       columnWidths: columnWidths,
//    },
//    subStyle: {
//       font: { bold: true, size: 10, color: { argb: '000000' } },
//       fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'f2f2f2' } },
//       border: { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} },
//       alignment: { horizontal: 'center', vertical: 'middle' },
//       rowHeight: 25,
//    },
//    bodyStyle: {
//       font: { size: 10, color: { argb: '000000' } },
//       fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } },
//       border: { top:{style:'thin'}, left:{style:'none'}, bottom:{style:'thin'}, right:{style:'dashed'} },
//       alignment: { horizontal: 'center', vertical: 'middle' },
//       rowHeight: 40,
//    },
// }

const EXCEL_DEFAULT_STYLE = {
   font:         { bold: true, size: 10, color: { argb: 'FFFFFFFF' } },
   fill:         { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } },
   border: {
      top:    { style: 'thin', color: { argb: 'FF000000' } },
      left:   { style: 'thin', color: { argb: 'FF000000' } },
      bottom: { style: 'thin', color: { argb: 'FF000000' } },
      right:  { style: 'thin', color: { argb: 'FF000000' } },
   },
   alignment:    { horizontal: 'center', vertical: 'middle', wrapText: true },
   columnWidths: 25,
};

const EXCEL_BODY_DEFAULT_STYLE = {
   ...EXCEL_DEFAULT_STYLE,
   font: { bold: false, size: 10, color: { argb: 'FF000000' } },
   fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } },
}

function _applyCellStyle(cell, style) {
   if (style.font)      cell.font      = style.font;
   if (style.fill)      cell.fill      = style.fill;
   if (style.border)    cell.border    = style.border;
   if (style.alignment) cell.alignment = style.alignment;
}

// CSS 클래스명 → 셀 스타일 덮어쓰기 (border/font/alignment 객체를 직접 변경)
function _applyClassToCell(cell, classStr, colorFromClass, styles) {
   if (!classStr) return;
   const { border, font, alignment } = styles;
   const classes = classStr.split(' ');
   const has = cls => classes.includes(cls);

   classes.forEach(cls => {
      const argb = colorFromClass[cls];
      if (argb) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb } };
   });

   if (has('left'))  alignment.horizontal = 'left';
   if (has('right')) alignment.horizontal = 'right';

   const borderMap = {
      topbold: 'top', leftbold: 'left', bottombold: 'bottom', rightbold: 'right',
      topdouble: 'top', leftdouble: 'left', bottomdouble: 'bottom', rightdouble: 'right',
   };
   Object.entries(borderMap).forEach(([cls, side]) => {
      if (has(cls)) border[side].style = cls.includes('double') ? 'double' : 'thin';
   });

   if (has('Saturday')) font.color.argb = 'FF0000FF';
   if (has('Sunday'))   font.color.argb = 'FFFF0000';
   if (has('fontbold')) font.bold = true;
}

function _writeHeaderRow(ws, rowData, rowIndex, style) {
   let startCol = 1, prevValue = null;
   rowData.forEach((item, i) => {
      const colIndex = i + 1;
      const cell = ws.getCell(rowIndex, colIndex);
      if (item !== 'NULL') {
         if (prevValue !== null && colIndex - 1 > startCol) {
            ws.mergeCells(rowIndex, startCol, rowIndex, colIndex - 1);
            applyBorderToRange(ws, rowIndex, startCol, rowIndex, colIndex - 1, style.border);
         }
         prevValue = item;
         startCol  = colIndex;
         if (typeof item === 'string' && item.includes('<br')) {
            cell.value     = item.replace(/<br\s*\/?>/gi, '\n');
            cell.alignment = { ...style.alignment, wrapText: true };
         } else {
            cell.value = item;
         }
         _applyCellStyle(cell, style);
      } else {
         cell.value = null;
      }
   });
   if (prevValue !== null && rowData.length >= startCol + 1) {
      ws.mergeCells(rowIndex, startCol, rowIndex, rowData.length);
      applyBorderToRange(ws, rowIndex, startCol, rowIndex, rowData.length, style.border);
   }
}

function _writeSubHeaderRow(ws, subRowData, rowIndex, mainStyle, subStyle) {
   let startCol = 1, prevValue = null;
   subRowData.forEach((item, i) => {
      const colIndex  = i + 1;
      const mainCell  = ws.getCell(rowIndex, colIndex);
      if (item === 'NULL2') {
         ws.mergeCells(rowIndex, colIndex, rowIndex + 1, colIndex);
         applyBorderToRange(ws, rowIndex, startCol, rowIndex, colIndex - 1, mainStyle.border);
         mainCell.value = mainCell.value ?? subRowData[i]; // 기존 값 없으면 원본값
         _applyCellStyle(mainCell, mainStyle);
         startCol = colIndex;
         prevValue = null;
      } else {
         const subCell = ws.getCell(rowIndex + 1, i + 1);
         _applyCellStyle(subCell, subStyle);
         if (item !== 'NULL') {
            if (prevValue !== null && colIndex - 1 > startCol) {
               ws.mergeCells(rowIndex + 1, startCol, rowIndex + 1, colIndex - 1);
               applyBorderToRange(ws, rowIndex + 1, startCol, rowIndex + 1, colIndex - 1, subStyle.border);
            }
            prevValue = item;
            startCol  = colIndex;
            subCell.value     = item;
            subCell.alignment = { ...subStyle.alignment, wrapText: true };
         } else {
            subCell.value = null;
         }
      }
   });
}

function _writeBodyRows(ws, bodyRowData, startRow, { mainStyle, bodyStyle, rowType, colType, colorFromClass }) {
   bodyRowData.forEach((row, rowIdx, rowArr) => {
      const excelRow = rowIdx + startRow;
      let mergeStart = 1, prevValue = null;

      row._excelsheet.forEach((item, colIdx, colArr) => {
         const colIndex = colIdx + 1;
         const cell = ws.getCell(excelRow, colIndex);

         if (item === 'NULL') { cell.value = null; return; }

         if (prevValue !== null && colIndex - 1 > mergeStart) {
            ws.mergeCells(excelRow, mergeStart, excelRow, colIndex - 1);
            applyBorderToRange(ws, excelRow, mergeStart, excelRow, colIndex - 1, mainStyle.border);
         }
         prevValue  = item;
         mergeStart = colIndex;

         // 값 채우기
         // if (item !== null && item !== '' && !isNaN(item)) {
         //    cell.value  = Number(item);
         //    cell.numFmt = '#,##0';
         // } else if (typeof item === 'string' && item.includes('<font')) {
         //    cell.value = { richText: parseHtmlToRichText(item, EXCEL_DEFAULT_STYLE.font.size) };
         // } else if (typeof item === 'string' && item.includes('<br')) {
         //    cell.value     = item.replace(/<br\s*\/?>/gi, '\n');
         //    cell.alignment = { ...cell.alignment, wrapText: true };
         // } else {
         //    cell.value = item;
         // }

         if( item === null ) item = 0;

         // if (item !== null && item !== "" && !isNaN(item)) {
         if (item === 0 || item === '0' || (item !== null && item !== '' && item !== undefined && !isNaN(item) && String(item).trim() !== '')) {
            cell.value = Number(item);
            cell.numFmt = "#,##0";
         } else if (typeof item === "string" && item.includes("<font")) {
            cell.value = { richText: parseHtmlToRichText(item, EXCEL_DEFAULT_STYLE.font.size) };
         } else if (typeof item === "string" && item.includes("<br")) {
            // HTML 줄바꿈을 Excel 줄바꿈으로 변환
            const text = item.replace(/<br\s*\/?>/gi, "\n");
            cell.value = text;
            // 줄바꿈 보이게 wrapText 설정
            cell.alignment = { ...cell.alignment, wrapText: true };
         } else {
            cell.value = item;
         }


         // 기본 스타일
         _applyCellStyle(cell, bodyStyle);

         // 교차 행 색상
         if (rowType.rowhighlight) {
            const isMatch = rowType.rowhighlight === 'odd'
               ? (rowIdx + 1) % 2 === 1
               : (rowIdx + 1) % 2 === 0;
            // if (isMatch)
            //    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEBEBEB' ?? 'FFFFFFFF' } };

            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: isMatch ? 'FFEBEBEB' : 'FFFFFFFF' } };
         }


         // 클래스 → 스타일 덮어쓰기 (objects mutated in-place)
         const border    = { top: {...bodyStyle.border.top}, left: {...bodyStyle.border.left}, bottom: {...bodyStyle.border.bottom}, right: {...bodyStyle.border.right} };
         const font      = { ...bodyStyle.font, color: { ...bodyStyle.font.color } };
         const alignment = { ...bodyStyle.alignment, wrapText: true };
         const classStr  = ((row._excelclass ?? row._class)?.[colIdx]) ?? '';
         _applyClassToCell(cell, classStr, colorFromClass, { border, font, alignment });

         // 첫 열 헤더 스타일
         if (colType.firstFlag && colIdx === 0) {
            cell.fill        = mainStyle.fill;
            font.color.argb  = mainStyle.font.color.argb;
            Object.assign(border, { ...mainStyle.border });
            border.right.style = 'thin';
            alignment.horizontal = mainStyle.alignment.horizontal;
         }

         // 마지막 열/행 실선 강제
         if (colIdx === colArr.length - 1) border.right.style  = 'thin';
         if (rowIdx === rowArr.length - 1) border.bottom.style = 'thin';

         cell.alignment = { ...alignment, wrapText: true };
         cell.font      = font;
         cell.border    = border;
      });

      if (bodyStyle.rowHeight) ws.getRow(excelRow).height = bodyStyle.rowHeight;
   });
}

function _doVerticalMerge(ws, startRow, endRow, colIdx) {
   if (endRow - startRow < 1) return;
   try {
      ws.mergeCells(startRow, colIdx, endRow, colIdx);
      const topCell = ws.getCell(startRow, colIdx);
      for (let r = startRow; r <= endRow; r++) {
         const cell  = ws.getCell(r, colIdx);
         cell.font      = topCell.font;
         cell.fill      = topCell.fill;
         cell.border    = topCell.border;
         cell.alignment = topCell.alignment;
      }
   } catch (e) {
      console.warn(`병합 오류 (${startRow}~${endRow}, col${colIdx}):`, e.message);
   }
}

function _applyVerticalMerges(ws, bodyRowData, startRow, totalCols) {
   for (let colIdx = 0; colIdx < totalCols; colIdx++) {
      let mergeStart = null;
      for (let rowIdx = 0; rowIdx < bodyRowData.length; rowIdx++) {
         const value      = bodyRowData[rowIdx]._excelsheet[colIdx];
         const excelRowIdx = rowIdx + startRow;
         if (value === 'NULL2') {
            if (mergeStart === null) mergeStart = excelRowIdx - 1;
         } else if (mergeStart !== null) {
            _doVerticalMerge(ws, mergeStart, excelRowIdx - 1, colIdx + 1);
            mergeStart = null;
         }
      }
      if (mergeStart !== null)
         _doVerticalMerge(ws, mergeStart, bodyRowData.length + startRow - 1, colIdx + 1);
   }
}


const _excelCellExport = (
   fileName, sheetName, rowIndex, rowType = {}, colType = { firstFlag: false },
   rowData, subRowData = null, bodyRowData, styleOptions = {}
) => {
   if (!bodyRowData?.length) { alert('데이터가 없습니다.'); return; }
   if (!rowData)             { console.log('제목(1행) 데이터가 없습니다.'); return; }

   try {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet(sheetName);

      const { mainStyle = {}, subStyle = {}, bodyStyle = {}, colorFromClass = {} } = styleOptions;
      const MainStyle = deepChangeOption(JSON.parse(JSON.stringify(EXCEL_DEFAULT_STYLE)), mainStyle);
      const SubStyle  = deepChangeOption(JSON.parse(JSON.stringify(EXCEL_DEFAULT_STYLE)), subStyle);
      const BodyStyle = deepChangeOption(JSON.parse(JSON.stringify(EXCEL_BODY_DEFAULT_STYLE)), bodyStyle);

      if (MainStyle.rowHeight) ws.getRow(rowIndex).height = MainStyle.rowHeight;
      if (Array.isArray(MainStyle.columnWidths))
         MainStyle.columnWidths.forEach((w, i) => (ws.getColumn(i + 1).width = w));

      _writeHeaderRow(ws, rowData, rowIndex, MainStyle);
      if (subRowData) _writeSubHeaderRow(ws, subRowData, rowIndex, MainStyle, SubStyle);

      const bodyStartRow = rowIndex + (subRowData ? 2 : 1);
      _writeBodyRows(ws, bodyRowData, bodyStartRow, {
         mainStyle: MainStyle, bodyStyle: BodyStyle, rowType, colType, colorFromClass,
      });
      _applyVerticalMerges(ws, bodyRowData, bodyStartRow, bodyRowData[0]._excelsheet.length);

      wb.xlsx.writeBuffer()
         .then(buffer => {
            const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            // .xls인 경우 application/vnd.ms-excel. 하지만 excelJS 는 xls지원 안한다고함.
               // xls로 해도 문제는 없으나, 엑셀에서 경고 얼럿이 뜰 수 있음.
            // .xlsx인 경우 application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
            const url  = URL.createObjectURL(blob); // blob URL 생성
            const link = Object.assign(document.createElement('a'), { href: url, download: fileName });
            document.body.appendChild(link); // DOM에 붙였다가
            link.click();
            document.body.removeChild(link); // 제거
            URL.revokeObjectURL(url); // 메모리 누수 방지
         })
         .catch(err => console.error('다운로드 실패', err));

   } catch (err) {
      console.error(err);
      alert(`데이터가 없습니다.\n데이터가 있음에도 오류가 지속된다면 관리자에게\n\n${location.href}\n\n해당 주소를 공유해주세요.`);
   }
};
