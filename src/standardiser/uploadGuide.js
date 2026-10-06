import assetSelection from '../assets/guides/asset-selection.png';
import propertyExport from '../assets/guides/property-export.jpg';
import tenantLeaseExport from '../assets/guides/tenant-lease-export.jpg';

var STEP_ONE = {
  title: 'Select assets and open Analysis',
  body: 'Select the required assets, then open Analysis. Refresh the view. Do not calculate.',
  src: assetSelection,
  alt: 'Navigator with the required assets selected and the Analysis tab highlighted.',
};

var GUIDES = {
  property: {
    steps: [
      STEP_ONE,
      {
        title: 'Refresh and export the Property file',
        body: 'Choose Property from the view list and refresh. Export as Microsoft Excel (XLSX). Calculation is not required.',
        src: propertyExport,
        alt: 'Analysis view set to Property, with Microsoft Excel (XLSX) highlighted in the export menu.',
      },
    ],
  },
  unit: {
    steps: [
      STEP_ONE,
      {
        title: 'Refresh and export the Tenant / Lease file',
        body: 'Choose Tenant/Lease from the view list and refresh. Export as Microsoft Excel (XLSX). Calculation is not required.',
        src: tenantLeaseExport,
        alt: 'Analysis view set to Tenant/Lease, with Microsoft Excel (XLSX) highlighted in the export menu.',
      },
    ],
  },
};

function canHover() {
  return window.matchMedia('(hover: hover) and (pointer: fine)').matches;
}

export function setupUploadGuides() {
  var triggers = Array.prototype.slice.call(document.querySelectorAll('[data-upload-guide]'));
  if (!triggers.length) return;

  var popover = document.createElement('div');
  popover.id = 'upload-guide';
  popover.className = 'upload-guide';
  popover.setAttribute('role', 'dialog');
  popover.setAttribute('aria-modal', 'false');
  popover.setAttribute('aria-labelledby', 'upload-guide-title');
  popover.hidden = true;
  popover.innerHTML =
    '<div class="upload-guide-arrow" aria-hidden="true"></div>' +
    '<header class="upload-guide-head">' +
      '<div>' +
        '<h2 id="upload-guide-title">How to prepare this file</h2>' +
        '<p>2 quick steps</p>' +
      '</div>' +
      '<button type="button" class="upload-guide-close" data-guide-close aria-label="Close guide">' +
        '<svg viewBox="0 0 16 16" width="14" height="14" aria-hidden="true" focusable="false">' +
          '<path d="M4 4l8 8M12 4l-8 8" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"></path>' +
        '</svg>' +
      '</button>' +
    '</header>' +
    '<ol class="upload-guide-steps"></ol>' +
    '<p class="upload-guide-foot"><span class="upload-guide-mark" aria-hidden="true"></span>Refresh only — calculation is not required.</p>';

  var lightbox = document.createElement('div');
  lightbox.className = 'upload-guide-lightbox';
  lightbox.hidden = true;
  lightbox.innerHTML =
    '<button type="button" class="upload-guide-lightbox-backdrop" data-lightbox-close aria-label="Close enlarged screenshot"></button>' +
    '<figure class="upload-guide-lightbox-frame">' +
      '<button type="button" class="upload-guide-lightbox-close" data-lightbox-close aria-label="Close enlarged screenshot">Close</button>' +
      '<img alt="">' +
      '<figcaption></figcaption>' +
    '</figure>';

  document.body.appendChild(popover);
  document.body.appendChild(lightbox);

  var stepsEl = popover.querySelector('.upload-guide-steps');
  var arrow = popover.querySelector('.upload-guide-arrow');
  var lightboxImg = lightbox.querySelector('img');
  var lightboxCaption = lightbox.querySelector('figcaption');

  var openKind = null;
  var pinned = false;
  var anchor = null;
  var showTimer = null;
  var hideTimer = null;
  var hoverCapable = canHover();

  function setExpanded(kind) {
    triggers.forEach(function (trigger) {
      trigger.setAttribute('aria-expanded', trigger.dataset.uploadGuide === kind ? 'true' : 'false');
    });
  }

  function render(kind) {
    var guide = GUIDES[kind];
    stepsEl.replaceChildren();
    guide.steps.forEach(function (step, index) {
      var item = document.createElement('li');
      var heading = document.createElement('h3');
      var number = document.createElement('span');
      number.className = 'upload-guide-num';
      number.textContent = String(index + 1);
      var title = document.createElement('span');
      title.textContent = step.title;
      heading.appendChild(number);
      heading.appendChild(title);

      var copy = document.createElement('p');
      copy.textContent = step.body;

      var shot = document.createElement('button');
      shot.type = 'button';
      shot.className = 'upload-guide-shot';
      shot.setAttribute('aria-label', 'Enlarge screenshot: ' + step.alt);
      var image = document.createElement('img');
      image.src = step.src;
      image.alt = step.alt;
      image.addEventListener('load', place);
      shot.appendChild(image);

      var hint = document.createElement('p');
      hint.className = 'upload-guide-enlarge';
      hint.textContent = 'Select to enlarge';

      item.appendChild(heading);
      item.appendChild(copy);
      item.appendChild(shot);
      item.appendChild(hint);
      stepsEl.appendChild(item);
    });
  }

  function place() {
    if (!anchor || popover.hidden) return;
    var margin = 12;
    var gap = 10;
    var sheet = window.innerWidth < 760;
    popover.classList.toggle('is-sheet', sheet);
    arrow.hidden = sheet;
    if (sheet) {
      popover.style.top = margin + 'px';
      popover.style.left = margin + 'px';
      popover.style.right = margin + 'px';
      popover.style.width = 'auto';
      popover.style.maxHeight = (window.innerHeight - margin * 2) + 'px';
      return;
    }

    popover.style.right = 'auto';
    popover.style.width = '680px';

    var rect = anchor.getBoundingClientRect();
    var spaceBelow = window.innerHeight - rect.bottom - gap - margin;
    var spaceAbove = rect.top - gap - margin;
    var openAbove = spaceBelow < 460 && spaceAbove > spaceBelow;
    var available = Math.max(240, openAbove ? spaceAbove : spaceBelow);
    popover.style.maxHeight = Math.min(740, available) + 'px';

    var width = popover.offsetWidth;
    var height = popover.offsetHeight;
    var top = openAbove ? rect.top - gap - height : rect.bottom + gap;
    if (openAbove && top < margin) top = margin;
    var left = rect.left;
    if (left + width > window.innerWidth - margin) left = window.innerWidth - margin - width;
    if (left < margin) left = margin;

    popover.style.top = Math.round(top) + 'px';
    popover.style.left = Math.round(left) + 'px';

    var arrowLeft = rect.left + rect.width / 2 - left - 6;
    arrowLeft = Math.max(16, Math.min(width - 28, arrowLeft));
    arrow.style.left = Math.round(arrowLeft) + 'px';
    arrow.classList.toggle('is-above', !openAbove);
    arrow.classList.toggle('is-below', openAbove);
  }

  function closeLightbox() {
    lightbox.hidden = true;
    lightboxImg.removeAttribute('src');
    lightboxImg.alt = '';
    lightboxCaption.textContent = '';
  }

  function close() {
    var returnTo = anchor;
    var focusInside = popover.contains(document.activeElement) || lightbox.contains(document.activeElement);
    clearTimeout(showTimer);
    clearTimeout(hideTimer);
    pinned = false;
    openKind = null;
    anchor = null;
    popover.hidden = true;
    closeLightbox();
    setExpanded(null);
    if (focusInside && returnTo) returnTo.focus();
  }

  function open(kind, nextAnchor, pin) {
    if (!GUIDES[kind]) return;
    clearTimeout(showTimer);
    clearTimeout(hideTimer);
    if (openKind !== kind) {
      render(kind);
      stepsEl.scrollTop = 0;
    }
    openKind = kind;
    anchor = nextAnchor;
    if (pin) pinned = true;
    popover.hidden = false;
    setExpanded(kind);
    place();
  }

  function scheduleShow(kind, nextAnchor) {
    if (!hoverCapable || pinned || window.innerWidth < 760) return;
    clearTimeout(hideTimer);
    clearTimeout(showTimer);
    showTimer = setTimeout(function () {
      if (!nextAnchor.matches(':hover') && !popover.matches(':hover')) return;
      open(kind, nextAnchor, false);
    }, popover.hidden ? 90 : 0);
  }

  function scheduleHide() {
    if (pinned || !lightbox.hidden) return;
    clearTimeout(hideTimer);
    hideTimer = setTimeout(function () {
      if (pinned || !lightbox.hidden) return;
      close();
    }, 160);
  }

  triggers.forEach(function (trigger) {
    var kind = trigger.dataset.uploadGuide;
    trigger.addEventListener('mouseenter', function () {
      scheduleShow(kind, trigger);
    });
    trigger.addEventListener('mouseleave', scheduleHide);
    trigger.addEventListener('keydown', function (event) {
      if (event.key === 'Enter' || event.key === ' ') event.stopPropagation();
    });
    trigger.addEventListener('click', function (event) {
      event.preventDefault();
      event.stopPropagation();
      if (!popover.hidden && openKind === kind && pinned) {
        close();
        return;
      }
      open(kind, trigger, true);
      if (event.detail === 0) popover.focus();
    });
  });

  popover.tabIndex = -1;
  popover.addEventListener('pointerdown', function () {
    pinned = true;
  });
  popover.addEventListener('mouseenter', function () {
    clearTimeout(hideTimer);
  });
  popover.addEventListener('mouseleave', scheduleHide);
  popover.addEventListener('click', function (event) {
    if (event.target.closest('[data-guide-close]')) {
      close();
      return;
    }
    var shot = event.target.closest('.upload-guide-shot');
    if (!shot) return;
    var image = shot.querySelector('img');
    if (!image) return;
    lightboxImg.src = image.src;
    lightboxImg.alt = image.alt;
    lightboxCaption.textContent = image.alt;
    lightbox.hidden = false;
    var lightboxClose = lightbox.querySelector('.upload-guide-lightbox-close');
    if (lightboxClose) lightboxClose.focus();
  });

  lightbox.addEventListener('click', function (event) {
    if (event.target.closest('[data-lightbox-close]')) closeLightbox();
  });

  document.addEventListener('pointerdown', function (event) {
    if (popover.hidden) return;
    var target = event.target;
    if (popover.contains(target) || lightbox.contains(target)) return;
    if (target.closest && target.closest('[data-upload-guide]')) return;
    close();
  });

  document.addEventListener('keydown', function (event) {
    if (event.key !== 'Escape') return;
    if (document.body.classList.contains('report-open') || document.body.classList.contains('config-open')) return;
    if (!lightbox.hidden) {
      event.stopPropagation();
      closeLightbox();
      return;
    }
    if (!popover.hidden) {
      event.stopPropagation();
      close();
    }
  });

  window.addEventListener('resize', place);
  window.addEventListener('scroll', place, true);
}
