/**
 * HTML Color Reference (HCR) - Modern Web Edition
 * Faithful recreation & modernization of the 1995 Visual Basic 4 application
 * Originally written by Christopher Fazendin
 */

(function () {
  'use strict';

  // State
  const state = {
    activeTarget: 'background', // 'background' | 'normal' | 'link' | 'vlink' | 'alink'
    outputFormat: 'html',       // 'html' | 'css'
    colors: {
      background: [192, 192, 192],
      normal: [0, 0, 0],
      link: [85, 0, 238],
      vlink: [85, 26, 139],
      alink: [85, 26, 139]
    }
  };

  // Presets definition
  const PRESETS = {
    classic: {
      name: '1995 Default (Silver)',
      background: [192, 192, 192],
      normal: [0, 0, 0],
      link: [85, 0, 238],
      vlink: [85, 26, 139],
      alink: [85, 26, 139]
    },
    win31: {
      name: 'Windows 3.1 Hot Dog Stand',
      background: [0, 0, 0],
      normal: [255, 255, 0],
      link: [255, 0, 0],
      vlink: [255, 140, 0],
      alink: [255, 255, 255]
    },
    geocities: {
      name: '90s GeoCities Neon',
      background: [0, 0, 128],
      normal: [255, 255, 0],
      link: [0, 255, 255],
      vlink: [255, 0, 255],
      alink: [255, 255, 255]
    },
    terminal: {
      name: 'Matrix Green Terminal',
      background: [13, 17, 23],
      normal: [0, 255, 102],
      link: [57, 255, 20],
      vlink: [0, 143, 17],
      alink: [255, 255, 255]
    },
    cyberpunk: {
      name: 'Cyberpunk 1997',
      background: [26, 0, 44],
      normal: [252, 238, 10],
      link: [0, 240, 255],
      vlink: [255, 0, 127],
      alink: [255, 255, 255]
    },
    paper: {
      name: 'Classic Paper & Ink',
      background: [251, 240, 217],
      normal: [43, 43, 43],
      link: [139, 0, 0],
      vlink: [75, 0, 130],
      alink: [178, 34, 34]
    },
    solarized: {
      name: 'Solarized Dark',
      background: [0, 43, 54],
      normal: [131, 148, 150],
      link: [38, 139, 210],
      vlink: [108, 113, 196],
      alink: [203, 75, 22]
    },
    oceanic: {
      name: 'Midnight Ocean',
      background: [15, 23, 42],
      normal: [226, 232, 240],
      link: [56, 189, 248],
      vlink: [168, 85, 247],
      alink: [244, 63, 94]
    }
  };

  // DOM Elements
  const el = {
    body: document.body,
    presetSelector: document.getElementById('presetSelector'),
    btnThemeWin95: document.getElementById('btnThemeWin95'),
    btnThemeModern: document.getElementById('btnThemeModern'),
    titleBarClose: document.getElementById('titleBarClose'),
    
    // Sliders & inputs
    hsbRed: document.getElementById('hsbRed'),
    hsbGreen: document.getElementById('hsbGreen'),
    hsbBlue: document.getElementById('hsbBlue'),
    txtRed: document.getElementById('txtRed'),
    txtGreen: document.getElementById('txtGreen'),
    txtBlue: document.getElementById('txtBlue'),
    
    // Swatch & picker
    activeColorPreview: document.getElementById('activeColorPreview'),
    activeHexBadge: document.getElementById('activeHexBadge'),
    nativeColorPicker: document.getElementById('nativeColorPicker'),
    btnEyeDropper: document.getElementById('btnEyeDropper'),
    
    // Preview Box
    picColor: document.getElementById('picColor'),
    lblNormal: document.getElementById('lblNormal'),
    lblLink: document.getElementById('lblLink'),
    lblVLink: document.getElementById('lblVLink'),
    lblALink: document.getElementById('lblALink'),
    
    // Contrast badges
    contrastNormal: document.getElementById('contrastNormal'),
    contrastLink: document.getElementById('contrastLink'),
    contrastVLink: document.getElementById('contrastVLink'),
    
    // Radio buttons
    radioBackground: document.getElementById('optBackground'),
    radioNormal: document.getElementById('optNormal'),
    radioLink: document.getElementById('optLink'),
    radioVLink: document.getElementById('optVLink'),
    radioALink: document.getElementById('optALink'),
    radioOptions: document.querySelectorAll('input[name="targetOption"]'),
    
    // Tag outputs
    bgTag: document.getElementById('bgTag'),
    txtTag: document.getElementById('txtTag'),
    linkTag: document.getElementById('linkTag'),
    vlinkTag: document.getElementById('vlinkTag'),
    alinkTag: document.getElementById('alinkTag'),
    txtMain: document.getElementById('txtMain'),
    btnCopyMain: document.getElementById('btnCopyMain'),
    
    // Format tabs
    tabHtml: document.getElementById('tabHtml'),
    tabCss: document.getElementById('tabCss'),
    
    // Buttons
    cmdAbout: document.getElementById('cmdAbout'),
    cmdExit: document.getElementById('cmdExit'),
    
    // Status Bar
    statusTarget: document.getElementById('statusTarget'),
    statusRgb: document.getElementById('statusRgb'),
    statusNotice: document.getElementById('statusNotice'),
    
    // Modal
    aboutModal: document.getElementById('aboutModal'),
    modalCloseBtn: document.getElementById('modalCloseBtn'),
    btnModalOk: document.getElementById('btnModalOk'),
    
    // Menu items
    menuResetDefaults: document.getElementById('menuResetDefaults'),
    menuCopyFullBody: document.getElementById('menuCopyFullBody'),
    menuCopyCSS: document.getElementById('menuCopyCSS'),
    menuExportHTML: document.getElementById('menuExportHTML'),
    menuCopyCurrentHex: document.getElementById('menuCopyCurrentHex'),
    menuRandomize: document.getElementById('menuRandomize'),
    menuInvert: document.getElementById('menuInvert'),
    menuToggleFormat: document.getElementById('menuToggleFormat'),
    menuToggleContrast: document.getElementById('menuToggleContrast'),
    menuAbout: document.getElementById('menuAbout'),
    
    // Toast
    toast: document.getElementById('toastNotification')
  };

  // Helper Functions
  function rgbToHex(r, g, b) {
    const toHex = (c) => {
      const hex = Math.max(0, Math.min(255, Math.round(c))).toString(16).toUpperCase();
      return hex.length === 1 ? '0' + hex : hex;
    };
    return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
  }

  function hexToRgb(hex) {
    let cleanHex = hex.replace(/^#/, '');
    if (cleanHex.length === 3) {
      cleanHex = cleanHex.split('').map(c => c + c).join('');
    }
    const num = parseInt(cleanHex, 16);
    if (isNaN(num) || cleanHex.length !== 6) return [0, 0, 0];
    return [(num >> 16) & 255, (num >> 8) & 255, num & 255];
  }

  function getLuminance(r, g, b) {
    const a = [r, g, b].map(v => {
      v /= 255;
      return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4);
    });
    return a[0] * 0.2126 + a[1] * 0.7152 + a[2] * 0.0722;
  }

  function getContrastRatio(rgb1, rgb2) {
    const lum1 = getLuminance(rgb1[0], rgb1[1], rgb1[2]);
    const lum2 = getLuminance(rgb2[0], rgb2[1], rgb2[2]);
    const brightest = Math.max(lum1, lum2);
    const darkest = Math.min(lum1, lum2);
    return (brightest + 0.05) / (darkest + 0.05);
  }

  function getContrastRating(ratio) {
    if (ratio >= 7) return { text: `${ratio.toFixed(1)}:1 (AAA)`, color: '#008800' };
    if (ratio >= 4.5) return { text: `${ratio.toFixed(1)}:1 (AA)`, color: '#006600' };
    if (ratio >= 3) return { text: `${ratio.toFixed(1)}:1 (Large AA)`, color: '#b8860b' };
    return { text: `${ratio.toFixed(1)}:1 (Fail)`, color: '#cc0000' };
  }

  function showToast(message) {
    if (!el.toast) return;
    el.toast.textContent = message;
    el.toast.classList.add('show');
    clearTimeout(el.toast._timer);
    el.toast._timer = setTimeout(() => {
      el.toast.classList.remove('show');
    }, 2000);
  }

  function copyToClipboard(text, label = 'Code') {
    navigator.clipboard.writeText(text).then(() => {
      showToast(`Copied ${label} to clipboard!`);
      if (el.statusNotice) el.statusNotice.textContent = 'Copied!';
      setTimeout(() => {
        if (el.statusNotice) el.statusNotice.textContent = 'Ready';
      }, 1500);
    }).catch(err => {
      showToast('Error copying to clipboard');
    });
  }

  // Update UI & Calculations
  function updateUI() {
    const currentRgb = state.colors[state.activeTarget];
    const r = currentRgb[0];
    const g = currentRgb[1];
    const b = currentRgb[2];
    const hex = rgbToHex(r, g, b);

    // Sync sliders and number boxes
    el.hsbRed.value = r;
    el.hsbGreen.value = g;
    el.hsbBlue.value = b;
    el.txtRed.value = r;
    el.txtGreen.value = g;
    el.txtBlue.value = b;

    // Swatch & picker
    el.activeColorPreview.style.backgroundColor = hex;
    el.activeHexBadge.textContent = hex;
    el.nativeColorPicker.value = hex.toLowerCase();

    // Update live preview elements
    const bgHex = rgbToHex(...state.colors.background);
    const txtHex = rgbToHex(...state.colors.normal);
    const linkHex = rgbToHex(...state.colors.link);
    const vlinkHex = rgbToHex(...state.colors.vlink);
    const alinkHex = rgbToHex(...state.colors.alink);

    el.picColor.style.backgroundColor = bgHex;
    el.lblNormal.style.color = txtHex;
    el.lblLink.style.color = linkHex;
    el.lblVLink.style.color = vlinkHex;
    el.lblALink.style.color = alinkHex;

    // Update preview target highlight
    document.querySelectorAll('.preview-element').forEach(item => {
      if (item.dataset.target === state.activeTarget) {
        item.classList.add('selected-target');
      } else {
        item.classList.remove('selected-target');
      }
    });

    // Update tag inputs
    el.bgTag.value = `bgcolor="${bgHex}"`;
    el.txtTag.value = `text="${txtHex}"`;
    el.linkTag.value = `link="${linkHex}"`;
    el.vlinkTag.value = `vlink="${vlinkHex}"`;
    el.alinkTag.value = `alink="${alinkHex}"`;

    // Update Main Output Tag / CSS
    if (state.outputFormat === 'html') {
      el.txtMain.value = `<body bgcolor="${bgHex}" text="${txtHex}" link="${linkHex}" vlink="${vlinkHex}" alink="${alinkHex}">`;
    } else {
      el.txtMain.value = `body {\n  background-color: ${bgHex};\n  color: ${txtHex};\n}\na:link {\n  color: ${linkHex};\n}\na:visited {\n  color: ${vlinkHex};\n}\na:active {\n  color: ${alinkHex};\n}`;
    }

    // Update Contrast Ratings
    const bgRgb = state.colors.background;
    const contrastNormal = getContrastRatio(bgRgb, state.colors.normal);
    const contrastLink = getContrastRatio(bgRgb, state.colors.link);
    const contrastVLink = getContrastRatio(bgRgb, state.colors.vlink);

    const rNormal = getContrastRating(contrastNormal);
    const rLink = getContrastRating(contrastLink);
    const rVLink = getContrastRating(contrastVLink);

    el.contrastNormal.innerHTML = `Text: <strong style="color:${rNormal.color}">${rNormal.text}</strong>`;
    el.contrastLink.innerHTML = `Link: <strong style="color:${rLink.color}">${rLink.text}</strong>`;
    el.contrastVLink.innerHTML = `Visited: <strong style="color:${rVLink.color}">${rVLink.text}</strong>`;

    // Update Status Bar
    const targetNames = {
      background: 'Background',
      normal: 'Normal Text',
      link: 'Link Text',
      vlink: 'Visited Link',
      alink: 'Active Link'
    };
    el.statusTarget.textContent = `Active: ${targetNames[state.activeTarget]} (${hex})`;
    el.statusRgb.textContent = `RGB(${r}, ${g}, ${b})`;
  }

  function setActiveTarget(target) {
    if (!state.colors[target]) return;
    state.activeTarget = target;

    // Sync radio buttons
    el.radioOptions.forEach(radio => {
      radio.checked = (radio.value === target);
    });

    updateUI();
  }

  function updateActiveColor(r, g, b) {
    state.colors[state.activeTarget] = [
      Math.max(0, Math.min(255, parseInt(r, 10) || 0)),
      Math.max(0, Math.min(255, parseInt(g, 10) || 0)),
      Math.max(0, Math.min(255, parseInt(g, 10) || 0))
    ];
    // Re-verify exact assignments
    state.colors[state.activeTarget][0] = Math.max(0, Math.min(255, parseInt(r, 10) || 0));
    state.colors[state.activeTarget][1] = Math.max(0, Math.min(255, parseInt(g, 10) || 0));
    state.colors[state.activeTarget][2] = Math.max(0, Math.min(255, parseInt(b, 10) || 0));
    updateUI();
  }

  function loadPreset(presetKey) {
    const p = PRESETS[presetKey];
    if (!p) return;
    state.colors.background = [...p.background];
    state.colors.normal = [...p.normal];
    state.colors.link = [...p.link];
    state.colors.vlink = [...p.vlink];
    state.colors.alink = [...p.alink];
    updateUI();
    showToast(`Loaded preset: ${p.name}`);
  }

  function randomizeColors() {
    const randomByte = () => Math.floor(Math.random() * 256);
    Object.keys(state.colors).forEach(key => {
      state.colors[key] = [randomByte(), randomByte(), randomByte()];
    });
    updateUI();
    showToast('Randomized color scheme!');
  }

  function invertColors() {
    Object.keys(state.colors).forEach(key => {
      state.colors[key] = state.colors[key].map(c => 255 - c);
    });
    updateUI();
    showToast('Inverted color scheme!');
  }

  function exportHTMLFile() {
    const bgHex = rgbToHex(...state.colors.background);
    const txtHex = rgbToHex(...state.colors.normal);
    const linkHex = rgbToHex(...state.colors.link);
    const vlinkHex = rgbToHex(...state.colors.vlink);
    const alinkHex = rgbToHex(...state.colors.alink);

    const htmlContent = `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>HTML Color Reference - Sample Page</title>
</head>
<body bgcolor="${bgHex}" text="${txtHex}" link="${linkHex}" vlink="${vlinkHex}" alink="${alinkHex}">
  <h1>Welcome to your HTML 2.0 / 3.2 Page!</h1>
  <p>This is standard normal text rendered with the text attribute.</p>
  <p><a href="#sample-link">This is a standard hypertext link</a></p>
  <p><a href="">This is a visited hypertext link</a></p>
  <hr>
  <p><small>Created with HTML Color Reference Web Edition (orig. 1995 by Christopher Fazendin)</small></p>
</body>
</html>`;

    const blob = new Blob([htmlContent], { type: 'text/html;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'color_sample.html';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    showToast('Exported sample HTML file!');
  }

  // Event Listeners
  function attachEventListeners() {
    // Sliders
    const handleSliderInput = () => {
      updateActiveColor(el.hsbRed.value, el.hsbGreen.value, el.hsbBlue.value);
    };
    el.hsbRed.addEventListener('input', handleSliderInput);
    el.hsbGreen.addEventListener('input', handleSliderInput);
    el.hsbBlue.addEventListener('input', handleSliderInput);

    // Number textboxes
    const handleNumberInput = () => {
      updateActiveColor(el.txtRed.value, el.txtGreen.value, el.txtBlue.value);
    };
    el.txtRed.addEventListener('input', handleNumberInput);
    el.txtGreen.addEventListener('input', handleNumberInput);
    el.txtBlue.addEventListener('input', handleNumberInput);

    // Native Color Picker
    el.nativeColorPicker.addEventListener('input', (e) => {
      const rgb = hexToRgb(e.target.value);
      updateActiveColor(rgb[0], rgb[1], rgb[2]);
    });

    // EyeDropper API
    el.btnEyeDropper.addEventListener('click', async () => {
      if ('EyeDropper' in window) {
        try {
          const eyeDropper = new window.EyeDropper();
          const result = await eyeDropper.open();
          const rgb = hexToRgb(result.sRGBHex);
          updateActiveColor(rgb[0], rgb[1], rgb[2]);
          showToast(`Picked ${result.sRGBHex}`);
        } catch (e) {
          // User cancelled dropper
        }
      } else {
        el.nativeColorPicker.click();
      }
    });

    // Radio button changes
    el.radioOptions.forEach(radio => {
      radio.addEventListener('change', (e) => {
        setActiveTarget(e.target.value);
      });
    });

    // Preview element clicks
    el.picColor.addEventListener('click', (e) => {
      if (e.target === el.picColor) {
        setActiveTarget('background');
      }
    });

    [el.lblNormal, el.lblLink, el.lblVLink, el.lblALink].forEach(item => {
      item.addEventListener('click', (e) => {
        e.stopPropagation();
        setActiveTarget(item.dataset.target);
      });
      item.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === ' ') {
          e.preventDefault();
          setActiveTarget(item.dataset.target);
        }
      });
    });

    // Copy buttons for individual tag inputs
    document.querySelectorAll('.btn-copy').forEach(btn => {
      btn.addEventListener('click', () => {
        const targetId = btn.dataset.copyTarget;
        const targetInput = document.getElementById(targetId);
        if (targetInput) {
          copyToClipboard(targetInput.value, targetInput.value);
        }
      });
    });

    // Copy full tag / CSS button
    el.btnCopyMain.addEventListener('click', () => {
      copyToClipboard(el.txtMain.value, state.outputFormat === 'html' ? '<body> tag' : 'CSS rules');
    });

    // Output Format tabs (HTML vs CSS)
    el.tabHtml.addEventListener('click', () => {
      state.outputFormat = 'html';
      el.tabHtml.classList.add('active');
      el.tabCss.classList.remove('active');
      el.btnCopyMain.textContent = 'Copy Tag';
      updateUI();
    });

    el.tabCss.addEventListener('click', () => {
      state.outputFormat = 'css';
      el.tabCss.classList.add('active');
      el.tabHtml.classList.remove('active');
      el.btnCopyMain.textContent = 'Copy CSS';
      updateUI();
    });

    // Preset selector
    el.presetSelector.addEventListener('change', (e) => {
      loadPreset(e.target.value);
    });

    // Theme Switchers
    el.btnThemeWin95.addEventListener('click', () => {
      el.body.className = 'theme-win95';
      el.btnThemeWin95.classList.add('active');
      el.btnThemeModern.classList.remove('active');
    });

    el.btnThemeModern.addEventListener('click', () => {
      el.body.className = 'theme-modern';
      el.btnThemeModern.classList.add('active');
      el.btnThemeWin95.classList.remove('active');
    });

    // About Modal
    const openAboutModal = () => {
      el.aboutModal.classList.add('open');
      el.aboutModal.setAttribute('aria-hidden', 'false');
      el.btnModalOk.focus();
    };

    const closeAboutModal = () => {
      el.aboutModal.classList.remove('open');
      el.aboutModal.setAttribute('aria-hidden', 'true');
    };

    el.cmdAbout.addEventListener('click', openAboutModal);
    el.menuAbout.addEventListener('click', openAboutModal);
    el.modalCloseBtn.addEventListener('click', closeAboutModal);
    el.btnModalOk.addEventListener('click', closeAboutModal);
    el.aboutModal.addEventListener('click', (e) => {
      if (e.target === el.aboutModal) closeAboutModal();
    });

    // Exit Button
    const handleExit = () => {
      if (confirm('Exit HTML Color Reference?')) {
        showToast('Application closed. You can refresh the page to reload.');
      }
    };
    el.cmdExit.addEventListener('click', handleExit);
    el.titleBarClose.addEventListener('click', handleExit);

    // Menu Bar actions
    el.menuResetDefaults.addEventListener('click', () => loadPreset('classic'));
    el.menuCopyFullBody.addEventListener('click', () => {
      const bgHex = rgbToHex(...state.colors.background);
      const txtHex = rgbToHex(...state.colors.normal);
      const linkHex = rgbToHex(...state.colors.link);
      const vlinkHex = rgbToHex(...state.colors.vlink);
      const alinkHex = rgbToHex(...state.colors.alink);
      copyToClipboard(`<body bgcolor="${bgHex}" text="${txtHex}" link="${linkHex}" vlink="${vlinkHex}" alink="${alinkHex}">`, '<body> tag');
    });
    el.menuCopyCSS.addEventListener('click', () => {
      const bgHex = rgbToHex(...state.colors.background);
      const txtHex = rgbToHex(...state.colors.normal);
      const linkHex = rgbToHex(...state.colors.link);
      const vlinkHex = rgbToHex(...state.colors.vlink);
      const alinkHex = rgbToHex(...state.colors.alink);
      const css = `body {\n  background-color: ${bgHex};\n  color: ${txtHex};\n}\na:link {\n  color: ${linkHex};\n}\na:visited {\n  color: ${vlinkHex};\n}\na:active {\n  color: ${alinkHex};\n}`;
      copyToClipboard(css, 'CSS');
    });
    el.menuExportHTML.addEventListener('click', exportHTMLFile);
    el.menuCopyCurrentHex.addEventListener('click', () => {
      const hex = rgbToHex(...state.colors[state.activeTarget]);
      copyToClipboard(hex, `Hex ${hex}`);
    });
    el.menuRandomize.addEventListener('click', randomizeColors);
    el.menuInvert.addEventListener('click', invertColors);
    el.menuToggleFormat.addEventListener('click', () => {
      if (state.outputFormat === 'html') el.tabCss.click();
      else el.tabHtml.click();
    });
    el.menuToggleContrast.addEventListener('click', () => {
      const panel = document.getElementById('contrastPanel');
      if (panel) panel.style.display = panel.style.display === 'none' ? 'flex' : 'none';
    });

    // Global Keyboard Shortcuts
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && el.aboutModal.classList.contains('open')) {
        closeAboutModal();
      }
    });
  }

  // Initialize
  function init() {
    attachEventListeners();
    updateUI();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
