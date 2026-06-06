<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
<title>Pelicana LA — Inventory</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
  @import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=DM+Sans:wght@400;500;600;700&display=swap');
  :root {
    --red:#D63B2F; --bg:#F7F5F2; --surface:#FFFFFF; --border:#E8E4DE;
    --text:#1A1614; --muted:#8A8480; --tag-bg:#EEEAE4; --green:#3DAB6A; --blue:#2F7DD6;
  }
  *{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent;}
  body{font-family:'DM Sans',sans-serif;background:var(--bg);color:var(--text);min-height:100vh;padding-bottom:120px;overscroll-behavior:none;}

  /* HEADER */
  .header{background:var(--red);color:white;padding:18px 20px 14px;position:sticky;top:0;z-index:100;box-shadow:0 2px 12px rgba(214,59,47,0.3);}
  .header-top{display:flex;align-items:center;justify-content:space-between;margin-bottom:12px;}
  .brand{font-size:11px;font-weight:600;letter-spacing:2px;text-transform:uppercase;opacity:.85;}
  .title{font-size:20px;font-weight:700;}
  .admin-icon{font-size:20px;cursor:pointer;opacity:.8;padding:4px;}
  .date-row{display:flex;align-items:center;gap:8px;background:rgba(255,255,255,0.15);border-radius:10px;padding:8px 12px;}
  .date-label{font-size:12px;opacity:.85;white-space:nowrap;}
  .date-input{background:transparent;border:none;color:white;font-family:'DM Mono',monospace;font-size:14px;font-weight:500;outline:none;flex:1;min-width:0;}
  .date-input::-webkit-calendar-picker-indicator{filter:invert(1);opacity:.7;}

  /* SETTINGS BAR */
  .settings-bar{background:var(--surface);border-bottom:1px solid var(--border);padding:10px 20px;display:flex;align-items:center;gap:8px;font-size:12px;color:var(--muted);}
  .settings-bar>span:first-child{flex:1;}
  .email-display{font-family:'DM Mono',monospace;font-size:12px;color:var(--text);font-weight:500;cursor:pointer;text-decoration:underline dotted;}

  /* TABS */
  .tabs{display:flex;background:var(--surface);border-bottom:1px solid var(--border);padding:0 20px;position:sticky;top:110px;z-index:90;}
  .tab{flex:1;padding:12px 8px;text-align:center;font-size:13px;font-weight:600;color:var(--muted);cursor:pointer;border-bottom:2.5px solid transparent;transition:all .2s;}
  .tab.active{color:var(--red);border-bottom-color:var(--red);}
  .view{display:none;}.view.active{display:block;}

  /* PROGRESS */
  .progress-bar{background:var(--surface);padding:10px 20px;border-bottom:1px solid var(--border);display:flex;align-items:center;gap:12px;}
  .progress-track{flex:1;height:4px;background:var(--border);border-radius:2px;overflow:hidden;}
  .progress-fill{height:100%;background:var(--red);border-radius:2px;transition:width .3s ease;}
  .progress-text{font-size:12px;color:var(--muted);font-family:'DM Mono',monospace;white-space:nowrap;}

  /* ITEMS */
  .category-section{margin:12px 16px 0;}
  .category-header{display:flex;align-items:center;gap:8px;padding:6px 0 8px;}
  .category-name{font-size:11px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:var(--muted);}
  .category-count{font-size:11px;font-family:'DM Mono',monospace;color:var(--muted);background:var(--tag-bg);padding:1px 7px;border-radius:20px;}
  .item-row{background:var(--surface);border-radius:12px;padding:10px 12px;margin-bottom:6px;display:flex;align-items:center;gap:10px;border:1.5px solid transparent;transition:border-color .15s;}
  .item-row.filled{border-color:#C6EAD5;background:#F5FFF9;}
  .item-info{flex:1;min-width:0;}
  .item-name{font-size:13px;font-weight:500;line-height:1.3;}
  .item-unit-row{display:flex;gap:4px;margin-top:2px;}
  .unit-pill{font-size:10px;font-family:'DM Mono',monospace;font-weight:600;padding:1px 6px;border-radius:20px;text-transform:uppercase;}
  .unit-pill.box{background:#E8F0FF;color:#2F5FBB;}
  .unit-pill.pack{background:#FFF3E0;color:#B86A00;}
  .badge-pair{display:flex;gap:6px;flex-shrink:0;}
  .badge-wrap{display:flex;flex-direction:column;align-items:center;gap:3px;}
  .badge-label{font-size:9px;font-weight:700;letter-spacing:.8px;color:var(--muted);text-transform:uppercase;}
  .qty-badge{min-width:46px;height:36px;border-radius:9px;display:flex;align-items:center;justify-content:center;font-family:'DM Mono',monospace;font-size:16px;font-weight:500;background:var(--bg);color:var(--muted);border:1.5px solid var(--border);cursor:pointer;transition:all .12s;user-select:none;}
  .qty-badge:active{opacity:.7;}
  .qty-badge.has-value.box-badge{background:#2F7DD6;color:white;border-color:#2F7DD6;}
  .qty-badge.has-value.pack-badge{background:#E07B00;color:white;border-color:#E07B00;}

  /* CAT TABS */
  .cat-tabs{display:flex;overflow-x:auto;padding:12px 16px 0;gap:8px;scrollbar-width:none;-webkit-overflow-scrolling:touch;}
  .cat-tabs::-webkit-scrollbar{display:none;}
  .cat-tab{flex-shrink:0;padding:7px 14px;border-radius:20px;font-size:12px;font-weight:600;background:var(--surface);border:1.5px solid var(--border);color:var(--muted);cursor:pointer;transition:all .15s;white-space:nowrap;}
  .cat-tab.active{background:var(--red);border-color:var(--red);color:white;}
  .cat-content{padding:8px 16px 0;}

  /* NUMPAD */
  .numpad-overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,0.6);z-index:300;align-items:flex-end;}
  .numpad-overlay.open{display:flex;}
  .numpad-sheet{background:#1C1C1E;border-radius:24px 24px 0 0;width:100%;max-width:480px;margin:0 auto;padding:0 0 32px;animation:slideUp .22s ease;}
  @keyframes slideUp{from{transform:translateY(100%)}to{transform:translateY(0)}}
  .numpad-item-info{padding:16px 20px 10px;border-bottom:1px solid rgba(255,255,255,0.08);}
  .numpad-item-cat{font-size:11px;font-weight:600;letter-spacing:1.5px;text-transform:uppercase;color:rgba(255,255,255,0.4);margin-bottom:2px;}
  .numpad-item-name{font-size:15px;font-weight:600;color:white;line-height:1.3;}
  .field-switcher{display:flex;gap:0;padding:12px 16px 0;}
  .field-btn{flex:1;padding:10px 8px;text-align:center;font-size:13px;font-weight:700;cursor:pointer;border:none;transition:all .15s;font-family:'DM Sans',sans-serif;}
  .field-btn:first-child{border-radius:10px 0 0 10px;}
  .field-btn:last-child{border-radius:0 10px 10px 0;}
  .field-btn.box-btn{background:rgba(47,125,214,0.15);color:rgba(255,255,255,0.4);}
  .field-btn.box-btn.active{background:#2F7DD6;color:white;}
  .field-btn.pack-btn{background:rgba(224,123,0,0.15);color:rgba(255,255,255,0.4);}
  .field-btn.pack-btn.active{background:#E07B00;color:white;}
  .numpad-display{padding:12px 24px 4px;display:flex;align-items:baseline;gap:8px;justify-content:flex-end;}
  .numpad-value{font-family:'DM Mono',monospace;font-size:52px;font-weight:500;color:white;line-height:1;min-width:60px;text-align:right;}
  .numpad-value.error{color:#FF453A;}
  .numpad-field-label{font-size:14px;color:rgba(255,255,255,0.35);font-family:'DM Mono',monospace;padding-bottom:8px;}
  .quick-row{display:flex;gap:6px;padding:4px 16px 10px;overflow-x:auto;scrollbar-width:none;}
  .quick-row::-webkit-scrollbar{display:none;}
  .quick-btn{flex-shrink:0;padding:6px 14px;border-radius:20px;font-size:14px;font-family:'DM Mono',monospace;font-weight:500;border:none;cursor:pointer;transition:background .1s;}
  .quick-btn.box-quick{background:rgba(47,125,214,0.25);color:#7EB8FF;}
  .quick-btn.pack-quick{background:rgba(224,123,0,0.25);color:#FFB84D;}
  .quick-btn:active{opacity:.7;}
  .keypad{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;padding:4px 16px 12px;}
  .key{height:60px;border-radius:12px;border:none;cursor:pointer;font-size:22px;font-family:'DM Sans',sans-serif;font-weight:600;transition:opacity .1s;user-select:none;display:flex;align-items:center;justify-content:center;}
  .key:active{opacity:.65;}
  .key-num{background:#2C2C2E;color:white;}
  .key-del{background:#2C2C2E;color:white;font-size:18px;}
  .key-clear{background:#3A3A3C;color:#FF9F0A;font-size:15px;font-weight:700;}
  .key-skip{background:#3A3A3C;color:rgba(255,255,255,0.4);font-size:13px;font-weight:600;}
  .key-next{grid-column:span 2;height:60px;border-radius:12px;border:none;font-size:15px;font-weight:700;font-family:'DM Sans',sans-serif;cursor:pointer;transition:opacity .1s;display:flex;align-items:center;justify-content:center;gap:6px;}
  .key-next.to-pack{background:#E07B00;color:white;}
  .key-next.confirm{background:var(--green);color:white;}
  .key-next:active{opacity:.85;}

  /* SEND BUTTON */
  .send-area{position:fixed;bottom:0;left:0;right:0;padding:12px 20px 28px;background:linear-gradient(to top,var(--bg) 70%,transparent);z-index:100;}
  .send-btn{width:100%;background:var(--red);color:white;border:none;border-radius:14px;padding:16px;font-size:16px;font-weight:700;font-family:'DM Sans',sans-serif;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:10px;box-shadow:0 4px 20px rgba(214,59,47,0.35);transition:transform .1s;}
  .send-btn:active{transform:scale(.98);}
  .send-count{background:rgba(255,255,255,0.25);border-radius:20px;padding:2px 10px;font-size:13px;}

  /* GENERIC MODAL BASE */
  .modal-overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,0.5);z-index:200;align-items:flex-end;}
  .modal-overlay.open{display:flex;}
  .modal{background:var(--surface);border-radius:20px 20px 0 0;padding:24px 20px 40px;width:100%;}
  .modal h3{font-size:16px;font-weight:700;margin-bottom:16px;}
  .modal input{width:100%;border:1.5px solid var(--border);border-radius:10px;padding:12px 14px;font-size:14px;font-family:'DM Mono',monospace;outline:none;margin-bottom:12px;}
  .modal input:focus{border-color:var(--red);}
  .modal-btns{display:flex;gap:10px;}
  .btn-cancel{flex:1;padding:12px;border:1.5px solid var(--border);background:none;border-radius:10px;font-size:14px;font-family:'DM Sans',sans-serif;cursor:pointer;}
  .btn-save{flex:1;padding:12px;background:var(--red);color:white;border:none;border-radius:10px;font-size:14px;font-family:'DM Sans',sans-serif;font-weight:600;cursor:pointer;}

  /* GUIDE MODAL */
  .guide-overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,0.6);z-index:500;align-items:flex-end;}
  .guide-overlay.open{display:flex;}
  .guide-sheet{background:var(--surface);border-radius:24px 24px 0 0;width:100%;padding:24px 20px 40px;animation:slideUp .22s ease;}
  .guide-title{font-size:18px;font-weight:700;margin-bottom:4px;}
  .guide-subtitle{font-size:13px;color:var(--muted);margin-bottom:20px;}
  .guide-steps{display:flex;flex-direction:column;gap:12px;margin-bottom:24px;}
  .guide-step{display:flex;align-items:flex-start;gap:14px;background:var(--bg);border-radius:14px;padding:14px;}
  .step-num{width:32px;height:32px;border-radius:50%;flex-shrink:0;display:flex;align-items:center;justify-content:center;font-size:14px;font-weight:700;color:white;}
  .step-num.done{background:var(--green);}
  .step-num.todo{background:var(--muted);}
  .step-num.action{background:var(--red);}
  .step-body{flex:1;}
  .step-label{font-size:13px;font-weight:700;margin-bottom:2px;}
  .step-desc{font-size:12px;color:var(--muted);line-height:1.5;}
  .step-file{margin-top:6px;background:white;border:1.5px solid var(--border);border-radius:8px;padding:8px 10px;font-family:'DM Mono',monospace;font-size:12px;color:#2F7DD6;word-break:break-all;}
  .guide-btns{display:flex;flex-direction:column;gap:10px;}
  .guide-open-mail{width:100%;padding:15px;background:var(--red);color:white;border:none;border-radius:14px;font-size:16px;font-weight:700;font-family:'DM Sans',sans-serif;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:8px;}
  .guide-close{width:100%;padding:13px;background:none;color:var(--muted);border:1.5px solid var(--border);border-radius:14px;font-size:14px;font-family:'DM Sans',sans-serif;cursor:pointer;}

  /* TOAST */
  .toast{position:fixed;bottom:110px;left:50%;transform:translateX(-50%) translateY(20px);background:#1A1614;color:white;padding:10px 20px;border-radius:20px;font-size:13px;opacity:0;transition:all .3s;pointer-events:none;white-space:nowrap;z-index:600;}
  .toast.show{opacity:1;transform:translateX(-50%) translateY(0);}

  /* ===== ADMIN ===== */
  /* PIN modal */
  .pin-overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,0.6);z-index:400;align-items:flex-end;}
  .pin-overlay.open{display:flex;}
  .pin-sheet{background:#1C1C1E;border-radius:24px 24px 0 0;width:100%;max-width:480px;margin:0 auto;padding:24px 20px 36px;animation:slideUp .22s ease;}
  .pin-title{font-size:16px;font-weight:700;color:white;text-align:center;margin-bottom:6px;}
  .pin-sub{font-size:12px;color:rgba(255,255,255,0.4);text-align:center;margin-bottom:20px;}
  .pin-dots{display:flex;justify-content:center;gap:14px;margin-bottom:20px;}
  .pin-dot{width:14px;height:14px;border-radius:50%;background:rgba(255,255,255,0.15);transition:background .15s;}
  .pin-dot.filled{background:white;}
  .pin-dot.error{background:#FF453A;}
  .pin-keypad{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;}
  .pin-key{height:64px;border-radius:14px;border:none;cursor:pointer;font-size:22px;font-family:'DM Sans',sans-serif;font-weight:600;background:#2C2C2E;color:white;display:flex;align-items:center;justify-content:center;transition:opacity .1s;}
  .pin-key:active{opacity:.6;}
  .pin-key.del{background:#3A3A3C;font-size:18px;}
  .pin-key.cancel{background:#3A3A3C;color:rgba(255,255,255,0.5);font-size:13px;font-weight:600;}

  /* Admin panel (full screen) */
  .admin-panel{display:none;position:fixed;inset:0;background:var(--bg);z-index:450;flex-direction:column;}
  .admin-panel.open{display:flex;}
  .admin-header{background:#1A1614;color:white;padding:18px 20px;display:flex;align-items:center;gap:12px;}
  .admin-back{font-size:22px;cursor:pointer;padding:2px 6px;}
  .admin-title{font-size:18px;font-weight:700;flex:1;}
  .admin-badge{background:var(--red);color:white;font-size:10px;font-weight:700;padding:3px 8px;border-radius:20px;letter-spacing:1px;}
  .admin-body{flex:1;overflow-y:auto;padding:16px;}

  /* Category accordion */
  .admin-cat{background:var(--surface);border-radius:14px;margin-bottom:10px;overflow:hidden;border:1.5px solid var(--border);}
  .admin-cat-header{display:flex;align-items:center;padding:14px 16px;cursor:pointer;gap:10px;}
  .admin-cat-name{font-size:14px;font-weight:700;flex:1;}
  .admin-cat-count{font-size:12px;font-family:'DM Mono',monospace;color:var(--muted);}
  .admin-cat-arrow{font-size:14px;color:var(--muted);transition:transform .2s;}
  .admin-cat.expanded .admin-cat-arrow{transform:rotate(90deg);}
  .admin-cat-items{display:none;border-top:1px solid var(--border);}
  .admin-cat.expanded .admin-cat-items{display:block;}

  .admin-item{display:flex;align-items:center;padding:10px 16px;gap:10px;border-bottom:1px solid var(--border);}
  .admin-item:last-child{border-bottom:none;}
  .admin-item-name{flex:1;font-size:13px;}
  .admin-item-unit{font-size:11px;font-family:'DM Mono',monospace;color:var(--muted);background:var(--tag-bg);padding:2px 8px;border-radius:20px;}
  .admin-item-edit{font-size:16px;cursor:pointer;padding:4px 6px;color:var(--blue);}
  .admin-item-del{font-size:16px;cursor:pointer;padding:4px 6px;color:var(--red);}

  .admin-add-item{display:flex;align-items:center;padding:10px 16px;gap:8px;background:var(--bg);}
  .admin-add-input{flex:1;border:1.5px solid var(--border);border-radius:8px;padding:8px 10px;font-size:13px;outline:none;font-family:'DM Sans',sans-serif;}
  .admin-add-input:focus{border-color:var(--blue);}
  .admin-unit-select{border:1.5px solid var(--border);border-radius:8px;padding:8px 6px;font-size:12px;font-family:'DM Mono',monospace;outline:none;background:white;color:var(--text);}
  .admin-add-btn{background:var(--blue);color:white;border:none;border-radius:8px;padding:8px 14px;font-size:13px;font-weight:700;font-family:'DM Sans',sans-serif;cursor:pointer;}

  .admin-add-cat-row{display:flex;gap:8px;margin-bottom:12px;}
  .admin-add-cat-input{flex:1;border:1.5px solid var(--border);border-radius:10px;padding:11px 14px;font-size:14px;outline:none;font-family:'DM Sans',sans-serif;}
  .admin-add-cat-input:focus{border-color:var(--blue);}
  .admin-add-cat-btn{background:var(--blue);color:white;border:none;border-radius:10px;padding:11px 18px;font-size:14px;font-weight:700;font-family:'DM Sans',sans-serif;cursor:pointer;}

  /* Edit item modal */
  .edit-overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,0.5);z-index:500;align-items:flex-end;}
  .edit-overlay.open{display:flex;}
  .edit-sheet{background:var(--surface);border-radius:20px 20px 0 0;padding:24px 20px 40px;width:100%;}
  .edit-sheet h3{font-size:16px;font-weight:700;margin-bottom:16px;}
  .edit-field{margin-bottom:12px;}
  .edit-label{font-size:11px;font-weight:700;color:var(--muted);letter-spacing:1px;text-transform:uppercase;margin-bottom:6px;}
  .edit-input{width:100%;border:1.5px solid var(--border);border-radius:10px;padding:12px 14px;font-size:14px;outline:none;font-family:'DM Sans',sans-serif;}
  .edit-input:focus{border-color:var(--blue);}
  .edit-select{width:100%;border:1.5px solid var(--border);border-radius:10px;padding:12px 14px;font-size:14px;outline:none;font-family:'DM Mono',monospace;background:white;color:var(--text);}
  .edit-btns{display:flex;gap:10px;margin-top:4px;}
</style>
</head>
<body>

<div class="header">
  <div class="header-top">
    <div>
      <div class="brand">Pelicana Chicken · LA</div>
      <div class="title">Weekly Inventory</div>
    </div>
    <div class="admin-icon" onclick="openPinModal()" title="관리자">⚙️</div>
  </div>
  <div class="date-row">
    <span class="date-label">📅 Date</span>
    <input type="date" class="date-input" id="invDate">
  </div>
</div>

<div class="settings-bar">
  <span>이메일 수신처</span>
  <span class="email-display" id="emailDisplay" onclick="openEmailModal()">fm.pelicanausa@gmail.com</span>
  <span onclick="openEmailModal()" style="cursor:pointer;">✏️</span>
</div>

<!-- EMAIL MODAL -->
<div class="modal-overlay" id="emailModal">
  <div class="modal">
    <h3>헤드쿼터 이메일 변경</h3>
    <input type="email" id="emailInput" placeholder="example@pelicana.com">
    <div class="modal-btns">
      <button class="btn-cancel" onclick="closeEmailModal()">취소</button>
      <button class="btn-save" onclick="saveEmail()">저장</button>
    </div>
  </div>
</div>

<div class="tabs">
  <div class="tab active" onclick="switchView('scroll')">전체 스크롤</div>
  <div class="tab" onclick="switchView('category')">카테고리별</div>
</div>
<div class="progress-bar">
  <div class="progress-track"><div class="progress-fill" id="progressFill" style="width:0%"></div></div>
  <div class="progress-text" id="progressText">0 / 0</div>
</div>

<div class="view active" id="view-scroll"><div id="scroll-content"></div></div>
<div class="view" id="view-category">
  <div class="cat-tabs" id="cat-tabs"></div>
  <div class="cat-content" id="cat-content"></div>
</div>

<!-- NUMPAD MODAL -->
<div class="numpad-overlay" id="numpadOverlay" onclick="handleOverlayClick(event)">
  <div class="numpad-sheet" id="numpadSheet">
    <div class="numpad-item-info">
      <div class="numpad-item-cat" id="np-cat"></div>
      <div class="numpad-item-name" id="np-name"></div>
    </div>
    <div class="field-switcher">
      <button class="field-btn box-btn active" id="btn-box" onclick="switchField('box')">📦 BOX</button>
      <button class="field-btn pack-btn" id="btn-pack" onclick="switchField('pack')">🗂 PACK / EA</button>
    </div>
    <div class="numpad-display">
      <div class="numpad-value" id="np-display">—</div>
      <div class="numpad-field-label" id="np-fieldlabel">BOX</div>
    </div>
    <div class="quick-row" id="quick-row"></div>
    <div class="keypad">
      <button class="key key-num" onclick="npKey('1')">1</button>
      <button class="key key-num" onclick="npKey('2')">2</button>
      <button class="key key-num" onclick="npKey('3')">3</button>
      <button class="key key-num" onclick="npKey('4')">4</button>
      <button class="key key-num" onclick="npKey('5')">5</button>
      <button class="key key-num" onclick="npKey('6')">6</button>
      <button class="key key-num" onclick="npKey('7')">7</button>
      <button class="key key-num" onclick="npKey('8')">8</button>
      <button class="key key-num" onclick="npKey('9')">9</button>
      <button class="key key-clear" onclick="npClear()">C</button>
      <button class="key key-num" onclick="npKey('0')">0</button>
      <button class="key key-del" onclick="npDel()">⌫</button>
      <button class="key key-skip" id="key-skip" onclick="npSkip()">건너뜀</button>
      <button class="key key-next confirm" id="key-next" onclick="npNext()">✓ 확인</button>
    </div>
  </div>
</div>

<!-- GUIDE MODAL -->
<div class="guide-overlay" id="guideOverlay">
  <div class="guide-sheet">
    <div class="guide-title">📤 인벤토리 전송 방법</div>
    <div class="guide-subtitle">아래 순서대로 따라하세요</div>
    <div class="guide-steps">
      <div class="guide-step">
        <div class="step-num done">✓</div>
        <div class="step-body">
          <div class="step-label">엑셀 파일 저장 완료</div>
          <div class="step-desc">다운로드 폴더에 저장되었습니다</div>
          <div class="step-file" id="guide-filename"></div>
        </div>
      </div>
      <div class="guide-step">
        <div class="step-num action">2</div>
        <div class="step-body">
          <div class="step-label">이메일 앱 열기</div>
          <div class="step-desc">아래 버튼으로 이메일 앱을 여세요.<br>수신자·제목·본문이 자동으로 입력돼 있어요.</div>
        </div>
      </div>
      <div class="guide-step">
        <div class="step-num todo">3</div>
        <div class="step-body">
          <div class="step-label">엑셀 파일 첨부 후 전송</div>
          <div class="step-desc">첨부 버튼(📎) → <b>다운로드 폴더</b>에서<br>저장된 파일을 선택 후 전송하세요.</div>
        </div>
      </div>
    </div>
    <div class="guide-btns">
      <button class="guide-open-mail" onclick="openMailApp()">✉️ 이메일 앱 열기</button>
      <button class="guide-close" onclick="closeGuide()">닫기</button>
    </div>
  </div>
</div>

<!-- PIN MODAL -->
<div class="pin-overlay" id="pinOverlay" onclick="handlePinOverlay(event)">
  <div class="pin-sheet" id="pinSheet">
    <div class="pin-title">🔒 관리자 모드</div>
    <div class="pin-sub">비밀번호 4자리를 입력하세요</div>
    <div class="pin-dots" id="pinDots">
      <div class="pin-dot" id="dot0"></div>
      <div class="pin-dot" id="dot1"></div>
      <div class="pin-dot" id="dot2"></div>
      <div class="pin-dot" id="dot3"></div>
    </div>
    <div class="pin-keypad">
      <button class="pin-key" onclick="pinKey('1')">1</button>
      <button class="pin-key" onclick="pinKey('2')">2</button>
      <button class="pin-key" onclick="pinKey('3')">3</button>
      <button class="pin-key" onclick="pinKey('4')">4</button>
      <button class="pin-key" onclick="pinKey('5')">5</button>
      <button class="pin-key" onclick="pinKey('6')">6</button>
      <button class="pin-key" onclick="pinKey('7')">7</button>
      <button class="pin-key" onclick="pinKey('8')">8</button>
      <button class="pin-key" onclick="pinKey('9')">9</button>
      <button class="pin-key cancel" onclick="closePinModal()">취소</button>
      <button class="pin-key" onclick="pinKey('0')">0</button>
      <button class="pin-key del" onclick="pinDel()">⌫</button>
    </div>
  </div>
</div>

<!-- ADMIN PANEL -->
<div class="admin-panel" id="adminPanel">
  <div class="admin-header">
    <span class="admin-back" onclick="closeAdmin()">←</span>
    <span class="admin-title">관리자 모드</span>
    <span class="admin-badge">ADMIN</span>
  </div>
  <div class="admin-body" id="adminBody"></div>
</div>

<!-- EDIT ITEM MODAL -->
<div class="edit-overlay" id="editOverlay">
  <div class="edit-sheet">
    <h3>아이템 수정</h3>
    <div class="edit-field">
      <div class="edit-label">이름</div>
      <input type="text" class="edit-input" id="editName" placeholder="아이템 이름">
    </div>
    <div class="edit-field">
      <div class="edit-label">단위</div>
      <select class="edit-select" id="editUnit">
        <option value="BOX">BOX</option>
        <option value="PACK">PACK</option>
        <option value="EA">EA</option>
        <option value="BAG">BAG</option>
        <option value="BOTTLE">BOTTLE</option>
        <option value="CAN">CAN</option>
      </select>
    </div>
    <div class="edit-btns">
      <button class="btn-cancel" onclick="closeEditModal()">취소</button>
      <button class="btn-save" onclick="saveEditItem()">저장</button>
    </div>
  </div>
</div>

<div class="send-area">
  <button class="send-btn" onclick="sendEmail()">
    <span>📤 이메일로 전송</span>
    <span class="send-count" id="sendCount">0개 입력됨</span>
  </button>
</div>

<div class="toast" id="toast"></div>

<script>
// ===== DATA =====
let ITEMS = [
  {cat:"Supply",name:"Chicken Box (S)",unit:"BOX"},
  {cat:"Supply",name:"Chicken Box (L)",unit:"BOX"},
  {cat:"Supply",name:"Logo 32 oz Plastic Bowl",unit:"PACK"},
  {cat:"Supply",name:"Logo 32 oz Plastic Bowl Lid",unit:"PACK"},
  {cat:"Supply",name:"Logo Side Box #1",unit:"BOX"},
  {cat:"Supply",name:"Logo To-Go Bag",unit:"BOX"},
  {cat:"Supply",name:"Cup Gangjung Box",unit:"PACK"},
  {cat:"Supply",name:"Cup Gangjung Box Lid",unit:"PACK"},
  {cat:"Supply",name:"Logo Napkin",unit:"BOX"},
  {cat:"Supply",name:"Logo Chopstick",unit:"BOX"},
  {cat:"Supply",name:"Logo Grease Paper (S)",unit:"BOX"},
  {cat:"Supply",name:"Logo Grease Paper (L)",unit:"BOX"},
  {cat:"Supply",name:"Paper Bag (Lunch)",unit:"PACK"},
  {cat:"Supply",name:"Produce Bag",unit:"BOX"},
  {cat:"Supply",name:"3.25 oz Portion Cup",unit:"PACK"},
  {cat:"Supply",name:"3.25 oz Portion Cup Lid",unit:"PACK"},
  {cat:"Supply",name:"To-Go Box #1 (leftover container) (SMALL)",unit:"BOX"},
  {cat:"Supply",name:"To-Go Box #2 (leftover container) (LARGE)",unit:"BOX"},
  {cat:"Supply",name:"2oz Cup",unit:"BOX"},
  {cat:"Supply",name:"2oz Cup Lid",unit:"BOX"},
  {cat:"Supply",name:"12 oz Paper Cup",unit:"PACK"},
  {cat:"Supply",name:"12 oz Paper Cup Lid",unit:"PACK"},
  {cat:"Supply",name:"14 oz PET Cup",unit:"BOX"},
  {cat:"Supply",name:"14 oz PET Cup Lid",unit:"PACK"},
  {cat:"Supply",name:"POS Roll Paper 3 1/8'' x 200'",unit:"BOX"},
  {cat:"Supply",name:"GP Nitrile Glove (L)",unit:"BOX"},
  {cat:"Supply",name:"Trash Bag 33 Gallon",unit:"BOX"},
  {cat:"Supply",name:"Trash Bag 55 Gallon",unit:"BOX"},
  {cat:"Supply",name:"Roll Paper Towel",unit:"BOX"},
  {cat:"Supply",name:"Toilet Paper",unit:"BOX"},
  {cat:"Supply",name:"Toilet Seat Cover",unit:"BOX"},
  {cat:"Supply",name:"Straw",unit:"PACK"},
  {cat:"Supply",name:"Fork",unit:"BOX"},
  {cat:"Sauce",name:"Signature Sauce",unit:"BOX"},
  {cat:"Sauce",name:"Hot Spicy Sauce",unit:"BOX"},
  {cat:"Sauce",name:"Smoky Hot Sauce",unit:"BOX"},
  {cat:"Sauce",name:"Smoky Mayo Sauce",unit:"BOX"},
  {cat:"Sauce",name:"Blazing Mayo Sauce",unit:"BOX"},
  {cat:"Sauce",name:"Garlic Sauce",unit:"BOX"},
  {cat:"Sauce",name:"White Garlic Sauce",unit:"BOX"},
  {cat:"Sauce",name:"Soy Sauce",unit:"BOX"},
  {cat:"Sauce",name:"Honey Soy Sauce",unit:"BOX"},
  {cat:"Sauce",name:"Gangjung Sauce",unit:"BOX"},
  {cat:"Sauce",name:"Buldak Sauce",unit:"BOX"},
  {cat:"Sauce",name:"Ddukbokki Sauce",unit:"BOX"},
  {cat:"Sauce",name:"Golbaengi Sauce",unit:"BOX"},
  {cat:"Sauce",name:"Fish Cake Soup Base",unit:"BOX"},
  {cat:"Sauce",name:"Signature Sauce To-Go",unit:"BOX"},
  {cat:"Sauce",name:"Hot Spicy Sauce To-Go",unit:"BOX"},
  {cat:"Sauce",name:"Smoky Hot Sauce To-Go",unit:"BOX"},
  {cat:"Sauce",name:"Smoky Mayo Sauce To-Go",unit:"BOX"},
  {cat:"Sauce",name:"Blazing Mayo Sauce To-Go",unit:"BOX"},
  {cat:"Sauce",name:"Garlic Sauce To-Go",unit:"BOX"},
  {cat:"Sauce",name:"Honey Soy Sauce To-Go",unit:"BOX"},
  {cat:"Powder",name:"Crispy Powder (Frying)",unit:"BOX"},
  {cat:"Powder",name:"Original Powder (Deep Frying)",unit:"BOX"},
  {cat:"Powder",name:"Marinade Powder",unit:"BOX"},
  {cat:"Powder",name:"Cheese Seasoning Powder",unit:"BOX"},
  {cat:"Powder",name:"Cinnamon Sugar for Twist Doughnut",unit:"PACK"},
  {cat:"Pickled Radish",name:"Pickled Radish Cube (S)",unit:"BOX"},
  {cat:"Pickled Radish",name:"Pickled Radish Cube (L)",unit:"BOX"},
  {cat:"Pickled Radish",name:"Pickled Radish Cube (Bulk)",unit:"BOX"},
  {cat:"Frozen",name:"Mozzar Cheese Ball",unit:"BOX"},
  {cat:"Frozen",name:"Twist Donut",unit:"BOX"},
  {cat:"Frozen",name:"Giant Seaweed Roll 대왕김말이",unit:"BOX"},
  {cat:"Frozen",name:"Undon Noodle 우동면",unit:"BOX"},
  {cat:"Frozen",name:"Fish Cake 사각어묵",unit:"BOX"},
  {cat:"Frozen",name:"Fish Cake Skewer 어묵꼬치",unit:"BOX"},
  {cat:"Frozen",name:"Spicy Pork Dumplings 매운 왕교자",unit:"BOX"},
  {cat:"Frozen",name:"Beef Dumplings 소고기 왕교자",unit:"BOX"},
  {cat:"Food",name:"Rice Cake 아씨 신당동 떡볶이 떡",unit:"BOX"},
  {cat:"Food",name:"Kimchi (10kg)",unit:"BOX"},
  {cat:"Food",name:"Seaweed flake 김가루 (500g)",unit:"PACK"},
  {cat:"Food",name:"Calrose Rice #40 Lbs",unit:"BAG"},
  {cat:"Food",name:"Sea Snail 유동 자연산 골뱅이",unit:"BOX"},
  {cat:"Food",name:"Spam",unit:"PACK"},
  {cat:"Food",name:"Sempio Wheat Noodles 샘표 진공 소면 (2.5KG*4)",unit:"BOX"},
  {cat:"Food",name:"Corn Syrup 물엿",unit:"BOX"},
  {cat:"Food",name:"MSG /Miwon 미원",unit:"PACK"},
  {cat:"Food",name:"Korean Chili powder 고운 고춧가루 (5 lbs)",unit:"PACK"},
  {cat:"Beverage",name:"CANOLA OIL",unit:"BOX"},
  {cat:"Beverage",name:"Jarritos - variety",unit:"BOX"},
  {cat:"Beverage",name:"Pellegrino",unit:"BOX"},
  {cat:"Beverage",name:"Coke",unit:"BOX"},
  {cat:"Beverage",name:"Sprite",unit:"BOX"},
  {cat:"Beverage",name:"Diet Coke",unit:"BOX"},
  {cat:"Alcohol",name:"Terra 테라",unit:"BOX"},
  {cat:"Alcohol",name:"Krush 크러쉬",unit:"BOX"},
  {cat:"Alcohol",name:"Fresh 참이슬 프레쉬",unit:"BOX"},
  {cat:"Alcohol",name:"Jinro 진로",unit:"BOX"},
  {cat:"Alcohol",name:"Jinro Strawberry 진로 딸기",unit:"BOX"},
  {cat:"Alcohol",name:"Jinro Green Grape 진로 청포도",unit:"BOX"},
  {cat:"Alcohol",name:"Zero Soju 새로소주",unit:"BOX"},
  {cat:"Alcohol",name:"Zero Soju Apricot 새로소주 살구",unit:"BOX"},
];

const ADMIN_PIN = '9999';
const QUICK_VALUES = [0,1,2,3,4,5,6,7,8,9,10];
const getCategories = () => [...new Set(ITEMS.map(i=>i.cat))];

let values = {};
let hqEmail = 'fm.pelicanausa@gmail.com';
let currentCat = getCategories()[0];
let npIdx=-1, npField='box', npBuffer='';
let savedMailtoLink='', savedFilename='';
let pinBuffer='';
let editingIdx=-1;

function resetValues() {
  values={};
  ITEMS.forEach((_,i)=>values[i]={box:'',pack:''});
}
resetValues();
document.getElementById('invDate').value = new Date().toISOString().slice(0,10);

// ===== PIN =====
function openPinModal() {
  pinBuffer='';
  renderPinDots();
  document.getElementById('pinOverlay').classList.add('open');
  document.body.style.overflow='hidden';
}
function closePinModal() {
  document.getElementById('pinOverlay').classList.remove('open');
  document.body.style.overflow='';
  pinBuffer='';
  renderPinDots();
}
function handlePinOverlay(e) { if(e.target===document.getElementById('pinOverlay')) closePinModal(); }
function pinKey(k) {
  if(pinBuffer.length>=4) return;
  pinBuffer+=k;
  renderPinDots();
  if(pinBuffer.length===4) setTimeout(checkPin,120);
}
function pinDel() { pinBuffer=pinBuffer.slice(0,-1); renderPinDots(); }
function renderPinDots(error=false) {
  for(let i=0;i<4;i++){
    const d=document.getElementById('dot'+i);
    d.classList.toggle('filled', i<pinBuffer.length);
    d.classList.toggle('error', error);
  }
}
function checkPin() {
  if(pinBuffer===ADMIN_PIN) {
    closePinModal();
    openAdmin();
  } else {
    renderPinDots(true);
    setTimeout(()=>{ pinBuffer=''; renderPinDots(); }, 600);
  }
}

// ===== ADMIN PANEL =====
function openAdmin() {
  buildAdminPanel();
  document.getElementById('adminPanel').classList.add('open');
  document.body.style.overflow='hidden';
}
function closeAdmin() {
  document.getElementById('adminPanel').classList.remove('open');
  document.body.style.overflow='';
  // Rebuild main views with updated items
  resetValues();
  buildScrollView();
  buildCategoryView();
  updateProgress();
}

function buildAdminPanel() {
  const body = document.getElementById('adminBody');
  const cats = getCategories();
  let html = `<div style="margin-bottom:16px;">
    <div style="font-size:11px;font-weight:700;color:var(--muted);letter-spacing:1.5px;text-transform:uppercase;margin-bottom:8px;">새 카테고리 추가</div>
    <div class="admin-add-cat-row">
      <input type="text" class="admin-add-cat-input" id="newCatInput" placeholder="카테고리 이름">
      <button class="admin-add-cat-btn" onclick="addCategory()">+ 추가</button>
    </div>
  </div>`;

  cats.forEach(cat => {
    const items = ITEMS.filter(i=>i.cat===cat);
    html += `<div class="admin-cat" id="acat-${css(cat)}">
      <div class="admin-cat-header" onclick="toggleAdminCat('${esc(cat)}')">
        <span class="admin-cat-name">${cat}</span>
        <span class="admin-cat-count">${items.length}개</span>
        <span class="admin-cat-arrow">▶</span>
      </div>
      <div class="admin-cat-items">`;
    items.forEach(item => {
      const idx = ITEMS.indexOf(item);
      html += `<div class="admin-item">
        <span class="admin-item-name">${item.name}</span>
        <span class="admin-item-unit">${item.unit}</span>
        <span class="admin-item-edit" onclick="openEditItem(${idx})">✏️</span>
        <span class="admin-item-del" onclick="deleteItem(${idx})">🗑</span>
      </div>`;
    });
    html += `<div class="admin-add-item">
      <input type="text" class="admin-add-input" id="newItem-${css(cat)}" placeholder="아이템 이름 추가">
      <select class="admin-unit-select" id="newUnit-${css(cat)}">
        <option>BOX</option><option>PACK</option><option>EA</option><option>BAG</option><option>BOTTLE</option><option>CAN</option>
      </select>
      <button class="admin-add-btn" onclick="addItem('${esc(cat)}')">+</button>
    </div>
    </div></div>`;
  });

  body.innerHTML = html;
}

function esc(s){ return s.replace(/'/g,"\\'"); }
function css(s){ return s.replace(/\s+/g,'_').replace(/[^a-zA-Z0-9_]/g,''); }

function toggleAdminCat(cat) {
  const el = document.getElementById('acat-'+css(cat));
  if(el) el.classList.toggle('expanded');
}

function addCategory() {
  const input = document.getElementById('newCatInput');
  const name = input.value.trim();
  if(!name){ showToast('카테고리 이름을 입력하세요'); return; }
  if(getCategories().includes(name)){ showToast('이미 있는 카테고리예요'); return; }
  ITEMS.push({cat:name, name:'(새 아이템)', unit:'BOX'});
  input.value='';
  buildAdminPanel();
  showToast(`"${name}" 카테고리 추가됨`);
}

function addItem(cat) {
  const nameEl = document.getElementById('newItem-'+css(cat));
  const unitEl = document.getElementById('newUnit-'+css(cat));
  const name = nameEl.value.trim();
  if(!name){ showToast('아이템 이름을 입력하세요'); return; }
  ITEMS.push({cat, name, unit: unitEl.value});
  nameEl.value='';
  buildAdminPanel();
  // Re-expand the category
  setTimeout(()=>{ const el=document.getElementById('acat-'+css(cat)); if(el) el.classList.add('expanded'); },50);
  showToast(`"${name}" 추가됨`);
}

function deleteItem(idx) {
  const item = ITEMS[idx];
  if(!confirm(`"${item.name}" 을(를) 삭제할까요?`)) return;
  ITEMS.splice(idx,1);
  buildAdminPanel();
  showToast('삭제되었습니다');
}

// Edit item modal
function openEditItem(idx) {
  editingIdx = idx;
  const item = ITEMS[idx];
  document.getElementById('editName').value = item.name;
  document.getElementById('editUnit').value = item.unit;
  document.getElementById('editOverlay').classList.add('open');
}
function closeEditModal() {
  document.getElementById('editOverlay').classList.remove('open');
  editingIdx=-1;
}
function saveEditItem() {
  if(editingIdx<0) return;
  const name = document.getElementById('editName').value.trim();
  const unit = document.getElementById('editUnit').value;
  if(!name){ showToast('이름을 입력하세요'); return; }
  ITEMS[editingIdx].name = name;
  ITEMS[editingIdx].unit = unit;
  closeEditModal();
  buildAdminPanel();
  showToast('수정되었습니다');
}

// ===== NUMPAD =====
function openNumpad(idx, startField) {
  npIdx=idx; npField=startField||'box';
  npBuffer = values[idx]&&values[idx][npField]!==''&&values[idx][npField]!==undefined ? String(values[idx][npField]) : '';
  const item=ITEMS[idx];
  document.getElementById('np-cat').textContent=item.cat;
  document.getElementById('np-name').textContent=item.name;
  syncFieldUI(); renderNpDisplay(); buildQuickRow();
  document.getElementById('numpadOverlay').classList.add('open');
  document.body.style.overflow='hidden';
}
function syncFieldUI() {
  const isBox=npField==='box';
  document.getElementById('btn-box').classList.toggle('active',isBox);
  document.getElementById('btn-pack').classList.toggle('active',!isBox);
  document.getElementById('np-fieldlabel').textContent=isBox?'BOX':'PACK / EA';
  const nb=document.getElementById('key-next');
  if(isBox){nb.textContent='다음 → PACK/EA';nb.className='key key-next to-pack';}
  else{nb.textContent='✓ 확인';nb.className='key key-next confirm';}
  buildQuickRow();
}
function switchField(field) { saveCurrentBuffer(); npField=field; npBuffer=values[npIdx][field]!==''?String(values[npIdx][field]):''; syncFieldUI(); renderNpDisplay(); }
function saveCurrentBuffer() {
  if(npIdx<0) return;
  const num=npBuffer===''?'':parseInt(npBuffer);
  if(npBuffer===''||(!isNaN(num)&&num<=9999)) { values[npIdx][npField]=npBuffer===''?'':num; refreshBadges(npIdx); }
}
function buildQuickRow() {
  const isBox=npField==='box';
  document.getElementById('quick-row').innerHTML=QUICK_VALUES.map(v=>`<button class="quick-btn ${isBox?'box-quick':'pack-quick'}" onclick="npQuick(${v})">${v}</button>`).join('');
}
function renderNpDisplay() {
  const el=document.getElementById('np-display');
  if(npBuffer===''){el.textContent='—';el.classList.remove('error');}
  else{el.textContent=npBuffer;el.classList.toggle('error',isNaN(parseInt(npBuffer))||parseInt(npBuffer)>9999);}
}
function npKey(k){if(npBuffer.length>=4)return;npBuffer+=k;renderNpDisplay();}
function npDel(){npBuffer=npBuffer.slice(0,-1);renderNpDisplay();}
function npClear(){npBuffer='';renderNpDisplay();}
function npQuick(v){
  npBuffer=String(v); renderNpDisplay();
  setTimeout(()=>{
    if(npField==='box'){saveCurrentBuffer();npField='pack';npBuffer=values[npIdx].pack!==''?String(values[npIdx].pack):'';syncFieldUI();renderNpDisplay();}
    else npNext();
  },130);
}
function npNext(){
  const num=npBuffer===''?'':parseInt(npBuffer);
  if(npBuffer!==''&&(isNaN(num)||num>9999)){showToast('올바른 숫자를 입력해주세요');return;}
  saveCurrentBuffer();
  if(npField==='box'){npField='pack';npBuffer=values[npIdx].pack!==''?String(values[npIdx].pack):'';syncFieldUI();renderNpDisplay();}
  else{refreshBadges(npIdx);updateProgress();closeNumpad();}
}
function npSkip(){saveCurrentBuffer();updateProgress();closeNumpad();}
function closeNumpad(){
  document.getElementById('numpadOverlay').classList.remove('open');
  document.body.style.overflow='';
  npIdx=-1;npBuffer='';
}
function handleOverlayClick(e){if(e.target===document.getElementById('numpadOverlay')){saveCurrentBuffer();updateProgress();closeNumpad();}}

// ===== BADGES / PROGRESS =====
function refreshBadges(idx) {
  const v=values[idx];
  const bf=v.box!==''&&v.box!==undefined, pf=v.pack!==''&&v.pack!==undefined;
  document.querySelectorAll(`[data-idx="${idx}"] .box-badge`).forEach(el=>{el.textContent=bf?v.box:'—';el.classList.toggle('has-value',bf);});
  document.querySelectorAll(`[data-idx="${idx}"] .pack-badge`).forEach(el=>{el.textContent=pf?v.pack:'—';el.classList.toggle('has-value',pf);});
  document.querySelectorAll(`[data-idx="${idx}"]`).forEach(el=>el.classList.toggle('filled',bf||pf));
}
function updateProgress() {
  const filled=Object.values(values).filter(v=>v.box!==''||v.pack!=='').length;
  const total=ITEMS.length;
  document.getElementById('progressFill').style.width=Math.round(filled/total*100)+'%';
  document.getElementById('progressText').textContent=`${filled} / ${total}`;
  document.getElementById('sendCount').textContent=`${filled}개 입력됨`;
}

// ===== BUILD UI =====
function makeItemRow(item,idx) {
  const v=values[idx]||{box:'',pack:''};
  const bf=v.box!==''&&v.box!==undefined, pf=v.pack!==''&&v.pack!==undefined;
  const div=document.createElement('div');
  div.className=`item-row${(bf||pf)?' filled':''}`;
  div.setAttribute('data-idx',idx);
  div.innerHTML=`
    <div class="item-info">
      <div class="item-name">${item.name}</div>
      <div class="item-unit-row"><span class="unit-pill box">BOX</span><span class="unit-pill pack">PACK/EA</span></div>
    </div>
    <div class="badge-pair">
      <div class="badge-wrap">
        <span class="badge-label">BOX</span>
        <div class="qty-badge box-badge${bf?' has-value':''}" onclick="openNumpad(${idx},'box')">${bf?v.box:'—'}</div>
      </div>
      <div class="badge-wrap">
        <span class="badge-label">PACK/EA</span>
        <div class="qty-badge pack-badge${pf?' has-value':''}" onclick="openNumpad(${idx},'pack')">${pf?v.pack:'—'}</div>
      </div>
    </div>`;
  return div;
}
function buildScrollView() {
  const container=document.getElementById('scroll-content');
  container.innerHTML='';
  getCategories().forEach(cat=>{
    const items=ITEMS.filter(i=>i.cat===cat);
    const sec=document.createElement('div');
    sec.className='category-section';
    sec.innerHTML=`<div class="category-header"><span class="category-name">${cat}</span><span class="category-count">${items.length}</span></div>`;
    items.forEach(item=>sec.appendChild(makeItemRow(item,ITEMS.indexOf(item))));
    container.appendChild(sec);
  });
}
function buildCategoryView() {
  currentCat=getCategories()[0]||'';
  document.getElementById('cat-tabs').innerHTML=getCategories().map(cat=>`<div class="cat-tab${cat===currentCat?' active':''}" onclick="selectCat('${esc(cat)}')">${cat}</div>`).join('');
  renderCatContent();
}
function renderCatContent() {
  const container=document.getElementById('cat-content');
  container.innerHTML='';
  const sec=document.createElement('div');
  sec.className='category-section';sec.style.margin='0';
  ITEMS.filter(i=>i.cat===currentCat).forEach(item=>sec.appendChild(makeItemRow(item,ITEMS.indexOf(item))));
  container.appendChild(sec);
}
function selectCat(cat) {
  currentCat=cat;
  document.querySelectorAll('.cat-tab').forEach(t=>t.classList.toggle('active',t.textContent===cat));
  renderCatContent();
}
function switchView(view) {
  document.querySelectorAll('.tab').forEach((t,i)=>t.classList.toggle('active',(view==='scroll'&&i===0)||(view==='category'&&i===1)));
  document.querySelectorAll('.view').forEach(v=>v.classList.remove('active'));
  document.getElementById(`view-${view}`).classList.add('active');
  if(view==='category') buildCategoryView();
}

// ===== EMAIL =====
function openEmailModal(){document.getElementById('emailInput').value=hqEmail;document.getElementById('emailModal').classList.add('open');}
function closeEmailModal(){document.getElementById('emailModal').classList.remove('open');}
function saveEmail(){const val=document.getElementById('emailInput').value.trim();if(val){hqEmail=val;document.getElementById('emailDisplay').textContent=hqEmail;showToast('이메일이 저장되었습니다');}closeEmailModal();}

// ===== GUIDE + SEND =====
function sendEmail() {
  const dateVal=document.getElementById('invDate').value;
  if(!dateVal){showToast('날짜를 선택해주세요');return;}
  const wb=XLSX.utils.book_new();
  const wsData=[['Pelicana LA Inventory'],['Date',dateVal],[],['Category','Description','BOX','PACK / EA','Selling Unit']];
  ITEMS.forEach((item,idx)=>{
    const v=values[idx]||{box:'',pack:''};
    wsData.push([item.cat,item.name,v.box!==''?Number(v.box):'',v.pack!==''?Number(v.pack):'',item.unit]);
  });
  const ws=XLSX.utils.aoa_to_sheet(wsData);
  ws['!cols']=[{wch:16},{wch:44},{wch:8},{wch:10},{wch:10}];
  XLSX.utils.book_append_sheet(wb,ws,'Inventory');
  savedFilename=`Pelicana_LA_Inventory_${dateVal}.xlsx`;
  const wbout = XLSX.write(wb, {bookType:'xlsx', type:'array'});
  const blob = new Blob([wbout], {type:'application/octet-stream'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = savedFilename;
  document.body.appendChild(a); a.click();
  document.body.removeChild(a);
  setTimeout(()=>URL.revokeObjectURL(url), 1000);
  const filled=Object.values(values).filter(v=>v.box!==''||v.pack!=='').length;
  const subject=encodeURIComponent(`[Pelicana LA] Weekly Inventory — ${dateVal}`);
  const body=encodeURIComponent(`안녕하세요,\n\nPelicana LA 주간 인벤토리 (${dateVal}) 파일을 첨부합니다.\n입력 항목: ${filled}개 / 전체 ${ITEMS.length}개\n\n감사합니다.`);
  savedMailtoLink=`mailto:${hqEmail}?subject=${subject}&body=${body}`;
  document.getElementById('guide-filename').textContent=savedFilename;
  document.getElementById('guideOverlay').classList.add('open');
  document.body.style.overflow='hidden';
}
function openMailApp(){window.location.href=savedMailtoLink;}
function closeGuide(){document.getElementById('guideOverlay').classList.remove('open');document.body.style.overflow='';}

// ===== TOAST =====
function showToast(msg){const t=document.getElementById('toast');t.textContent=msg;t.classList.add('show');setTimeout(()=>t.classList.remove('show'),2500);}

// INIT
buildScrollView();
updateProgress();
</script>
</body>
</html>
