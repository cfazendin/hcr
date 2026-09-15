/**
 * HTML Color Reference v2.04
 * Faithful 1:1 JavaScript implementation of frmHCR.frm & frmAbout.frm (Visual Basic 4, 1995)
 * Original by Christopher Fazendin
 */

(function () {
  'use strict';

  // State matching original VB4 variables:
  // SelOption: 1 = Background, 2 = Normal, 3 = Link, 4 = VLink, 5 = ALink
  let SelOption = 1;
  const BackGrnd = [192, 192, 192];
  const NormalTxt = [0, 0, 0];
  const LinkTxt = [85, 0, 238];
  const VLinkTxt = [85, 26, 139];
  const ALinkTxt = [85, 26, 139];

  // DOM Elements
  const el = {
    // Sliders & Textboxes
    hsbRed: document.getElementById('hsbRed'),
    hsbGreen: document.getElementById('hsbGreen'),
    hsbBlue: document.getElementById('hsbBlue'),
    txtRed: document.getElementById('txtRed'),
    txtGreen: document.getElementById('txtGreen'),
    txtBlue: document.getElementById('txtBlue'),

    // Options (Radio Buttons)
    optBackground: document.getElementById('optBackground'),
    optNormal: document.getElementById('optNormal'),
    optLink: document.getElementById('optLink'),
    optVLink: document.getElementById('optVLink'),
    optALink: document.getElementById('optALink'),

    // PictureBox & Labels
    picColor: document.getElementById('picColor'),
    lblNormal: document.getElementById('lblNormal'),
    lblLink: document.getElementById('lblLink'),
    lblVLink: document.getElementById('lblVLink'),
    lblALink: document.getElementById('lblALink'),

    // Tag TextBoxes
    bgTag: document.getElementById('bgTag'),
    txtTag: document.getElementById('txtTag'),
    linkTag: document.getElementById('linkTag'),
    vlinkTag: document.getElementById('vlinkTag'),
    alinkTag: document.getElementById('alinkTag'),
    txtMain: document.getElementById('txtMain'),

    // Buttons
    cmdAbout: document.getElementById('cmdAbout'),
    cmdExit: document.getElementById('cmdExit'),
    titleBarClose: document.getElementById('titleBarClose'),

    // Menus
    menuExit: document.getElementById('menuExit'),
    menuCopy: document.getElementById('menuCopy'),
    menuAbout: document.getElementById('menuAbout'),

    // Modal
    aboutModal: document.getElementById('aboutModal'),
    modalCloseBtn: document.getElementById('modalCloseBtn'),
    btnModalOk: document.getElementById('btnModalOk'),

    // Web Safe Palette Grid
    webSafePaletteGrid: document.getElementById('webSafePaletteGrid'),

    // Toast
    toast: document.getElementById('toastNotification')
  };

  function toHex2(val) {
    const hex = Math.max(0, Math.min(255, Math.round(val))).toString(16).toUpperCase();
    return hex.length === 1 ? '0' + hex : hex;
  }

  function hexToRgb(hex) {
    let clean = hex.replace(/^#/, '');
    if (clean.length === 3) {
      clean = clean.split('').map(c => c + c).join('');
    }
    const num = parseInt(clean, 16);
    if (isNaN(num)) return [0, 0, 0];
    return [(num >> 16) & 255, (num >> 8) & 255, num & 255];
  }

  function showToast(msg) {
    if (!el.toast) return;
    el.toast.textContent = msg;
    el.toast.classList.add('show');
    clearTimeout(el.toast._timer);
    el.toast._timer = setTimeout(() => el.toast.classList.remove('show'), 1800);
  }

  // Faithful ShowHex() from frmHCR.frm
  function ShowHex() {
    const bgHex = toHex2(BackGrnd[0]) + toHex2(BackGrnd[1]) + toHex2(BackGrnd[2]);
    const txtHex = toHex2(NormalTxt[0]) + toHex2(NormalTxt[1]) + toHex2(NormalTxt[2]);
    const linkHex = toHex2(LinkTxt[0]) + toHex2(LinkTxt[1]) + toHex2(LinkTxt[2]);
    const vlinkHex = toHex2(VLinkTxt[0]) + toHex2(VLinkTxt[1]) + toHex2(VLinkTxt[2]);
    const alinkHex = toHex2(ALinkTxt[0]) + toHex2(ALinkTxt[1]) + toHex2(ALinkTxt[2]);

    el.bgTag.value = `bgcolor="#${bgHex}"`;
    el.txtTag.value = `text="#${txtHex}"`;
    el.linkTag.value = `link="#${linkHex}"`;
    el.vlinkTag.value = `vlink="#${vlinkHex}"`;
    el.alinkTag.value = `alink="#${alinkHex}"`;

    el.txtMain.value = `<body ${el.bgTag.value} ${el.txtTag.value} ${el.linkTag.value} ${el.vlinkTag.value} ${el.alinkTag.value}>`;
  }

  // Faithful FindOption() from frmHCR.frm
  function FindOption() {
    if (el.optBackground.checked) {
      if (SelOption !== 1) {
        el.hsbRed.value = BackGrnd[0];
        el.hsbGreen.value = BackGrnd[1];
        el.hsbBlue.value = BackGrnd[2];
      }
      el.txtRed.value = el.hsbRed.value;
      el.txtGreen.value = el.hsbGreen.value;
      el.txtBlue.value = el.hsbBlue.value;
      BackGrnd[0] = parseInt(el.hsbRed.value, 10);
      BackGrnd[1] = parseInt(el.hsbGreen.value, 10);
      BackGrnd[2] = parseInt(el.hsbBlue.value, 10);
      SelOption = 1;
    } else if (el.optNormal.checked) {
      if (SelOption !== 2) {
        el.hsbRed.value = NormalTxt[0];
        el.hsbGreen.value = NormalTxt[1];
        el.hsbBlue.value = NormalTxt[2];
      }
      el.txtRed.value = el.hsbRed.value;
      el.txtGreen.value = el.hsbGreen.value;
      el.txtBlue.value = el.hsbBlue.value;
      NormalTxt[0] = parseInt(el.hsbRed.value, 10);
      NormalTxt[1] = parseInt(el.hsbGreen.value, 10);
      NormalTxt[2] = parseInt(el.hsbBlue.value, 10);
      SelOption = 2;
    } else if (el.optLink.checked) {
      if (SelOption !== 3) {
        el.hsbRed.value = LinkTxt[0];
        el.hsbGreen.value = LinkTxt[1];
        el.hsbBlue.value = LinkTxt[2];
      }
      el.txtRed.value = el.hsbRed.value;
      el.txtGreen.value = el.hsbGreen.value;
      el.txtBlue.value = el.hsbBlue.value;
      LinkTxt[0] = parseInt(el.hsbRed.value, 10);
      LinkTxt[1] = parseInt(el.hsbGreen.value, 10);
      LinkTxt[2] = parseInt(el.hsbBlue.value, 10);
      SelOption = 3;
    } else if (el.optVLink.checked) {
      if (SelOption !== 4) {
        el.hsbRed.value = VLinkTxt[0];
        el.hsbGreen.value = VLinkTxt[1];
        el.hsbBlue.value = VLinkTxt[2];
      }
      el.txtRed.value = el.hsbRed.value;
      el.txtGreen.value = el.hsbGreen.value;
      el.txtBlue.value = el.hsbBlue.value;
      VLinkTxt[0] = parseInt(el.hsbRed.value, 10);
      VLinkTxt[1] = parseInt(el.hsbGreen.value, 10);
      VLinkTxt[2] = parseInt(el.hsbBlue.value, 10);
      SelOption = 4;
    } else {
      if (SelOption !== 5) {
        el.hsbRed.value = ALinkTxt[0];
        el.hsbGreen.value = ALinkTxt[1];
        el.hsbBlue.value = ALinkTxt[2];
      }
      el.txtRed.value = el.hsbRed.value;
      el.txtGreen.value = el.hsbGreen.value;
      el.txtBlue.value = el.hsbBlue.value;
      ALinkTxt[0] = parseInt(el.hsbRed.value, 10);
      ALinkTxt[1] = parseInt(el.hsbGreen.value, 10);
      ALinkTxt[2] = parseInt(el.hsbBlue.value, 10);
      SelOption = 5;
    }
  }

  // Faithful FindColor() from frmHCR.frm
  function FindColor() {
    FindOption();
    const r = parseInt(el.hsbRed.value, 10);
    const g = parseInt(el.hsbGreen.value, 10);
    const b = parseInt(el.hsbBlue.value, 10);
    const rgbStr = `rgb(${r}, ${g}, ${b})`;

    switch (SelOption) {
      case 1:
        el.picColor.style.backgroundColor = rgbStr;
        break;
      case 2:
        el.lblNormal.style.color = rgbStr;
        break;
      case 3:
        el.lblLink.style.color = rgbStr;
        break;
      case 4:
        el.lblVLink.style.color = rgbStr;
        break;
      case 5:
        el.lblALink.style.color = rgbStr;
        break;
    }

    ShowHex();
  }

  // Apply a color directly to active target (used by 216 palette)
  function applyColorToActiveTarget(hex) {
    const rgb = hexToRgb(hex);
    el.hsbRed.value = rgb[0];
    el.hsbGreen.value = rgb[1];
    el.hsbBlue.value = rgb[2];
    FindColor();
    showToast(`Loaded ${hex}`);
  }

  // Render 216 Netscape Web-Safe Palette
  function renderWebSafePalette() {
    if (!el.webSafePaletteGrid) return;
    el.webSafePaletteGrid.innerHTML = '';
    const steps = ['00', '33', '66', '99', 'CC', 'FF'];

    for (let r = 0; r < 6; r++) {
      for (let g = 0; g < 6; g++) {
        for (let b = 0; b < 6; b++) {
          const hex = `#${steps[r]}${steps[g]}${steps[b]}`;
          const btn = document.createElement('button');
          btn.className = 'websafe-swatch';
          btn.style.backgroundColor = hex;
          btn.title = `Web-Safe: ${hex}`;
          btn.setAttribute('aria-label', `Web-Safe ${hex}`);
          btn.addEventListener('click', (e) => {
            e.preventDefault();
            applyColorToActiveTarget(hex);
          });
          el.webSafePaletteGrid.appendChild(btn);
        }
      }
    }
  }

  // Event Listeners
  function attachEventListeners() {
    // Sliders
    el.hsbRed.addEventListener('input', () => FindColor());
    el.hsbGreen.addEventListener('input', () => FindColor());
    el.hsbBlue.addEventListener('input', () => FindColor());

    // Textboxes Keypress / Enter
    const handleTxtInput = (txtInput, slider) => {
      let val = parseInt(txtInput.value.replace(/[^0-9]/g, ''), 10);
      if (isNaN(val)) val = 0;
      val = Math.max(0, Math.min(255, val));
      slider.value = val;
      FindColor();
    };

    el.txtRed.addEventListener('input', () => handleTxtInput(el.txtRed, el.hsbRed));
    el.txtGreen.addEventListener('input', () => handleTxtInput(el.txtGreen, el.hsbGreen));
    el.txtBlue.addEventListener('input', () => handleTxtInput(el.txtBlue, el.hsbBlue));

    // Radio Options Click
    [el.optBackground, el.optNormal, el.optLink, el.optVLink, el.optALink].forEach(opt => {
      opt.addEventListener('change', () => FindOption());
    });

    // PictureBox & Labels Click (faithful to VB4 events)
    el.picColor.addEventListener('click', (e) => {
      if (e.target === el.picColor) {
        el.optBackground.checked = true;
        FindOption();
      }
    });

    el.lblNormal.addEventListener('click', (e) => {
      e.stopPropagation();
      el.optNormal.checked = true;
      FindOption();
    });

    el.lblLink.addEventListener('click', (e) => {
      e.stopPropagation();
      el.optLink.checked = true;
      FindOption();
    });

    el.lblVLink.addEventListener('click', (e) => {
      e.stopPropagation();
      el.optVLink.checked = true;
      FindOption();
    });

    el.lblALink.addEventListener('click', (e) => {
      e.stopPropagation();
      el.optALink.checked = true;
      FindOption();
    });

    // Auto-select text on click for tag boxes
    [el.bgTag, el.txtTag, el.linkTag, el.vlinkTag, el.alinkTag, el.txtMain].forEach(box => {
      box.addEventListener('click', () => {
        box.select();
        navigator.clipboard.writeText(box.value).then(() => {
          showToast(`Copied: ${box.value}`);
        }).catch(() => {});
      });
    });

    // About Modal
    const openAbout = () => {
      el.aboutModal.classList.add('open');
      el.btnModalOk.focus();
    };
    const closeAbout = () => el.aboutModal.classList.remove('open');

    el.cmdAbout.addEventListener('click', openAbout);
    el.menuAbout.addEventListener('click', openAbout);
    el.modalCloseBtn.addEventListener('click', closeAbout);
    el.btnModalOk.addEventListener('click', closeAbout);
    el.aboutModal.addEventListener('click', (e) => {
      if (e.target === el.aboutModal) closeAbout();
    });

    // Exit
    const handleExit = () => {
      if (confirm('Exit HTML Color Reference?')) {
        window.close();
      }
    };
    el.cmdExit.addEventListener('click', handleExit);
    el.menuExit.addEventListener('click', handleExit);
    el.titleBarClose.addEventListener('click', handleExit);

    // Menu Copy
    el.menuCopy.addEventListener('click', () => {
      const active = document.activeElement;
      if (active && (active.tagName === 'INPUT' || active.tagName === 'TEXTAREA')) {
        document.execCommand('copy');
        showToast('Copied to clipboard');
      } else {
        navigator.clipboard.writeText(el.txtMain.value).then(() => {
          showToast('Copied <body ...> tag');
        });
      }
    });

    // Keyboard Shortcuts
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && el.aboutModal.classList.contains('open')) {
        closeAbout();
      }
    });
  }

  // Init
  function init() {
    renderWebSafePalette();
    attachEventListeners();
    FindOption();
    FindColor();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
