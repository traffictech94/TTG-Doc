(() => {
  const storageKey = 'ttg-theme';
  const root = document.documentElement;

  function readTheme() {
    try {
      return localStorage.getItem(storageKey) === 'dark' ? 'dark' : 'light';
    } catch {
      return 'light';
    }
  }

  function saveTheme(theme) {
    root.dataset.theme = theme;
    try {
      localStorage.setItem(storageKey, theme);
    } catch {
      // The current page still changes theme when browser storage is disabled.
    }
    updateToggle();
  }

  function updateToggle() {
    const button = document.getElementById('themeToggle');
    if (!button) return;
    button.querySelector('.theme-toggle-label')?.remove();
    const dark = root.dataset.theme === 'dark';
    button.setAttribute('aria-label', dark ? 'Switch to light mode' : 'Switch to dark mode');
    button.title = dark ? 'Switch to light mode' : 'Switch to dark mode';
    button.setAttribute('aria-pressed', String(dark));
    button.innerHTML = dark
      ? '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="4"/><path d="M12 2v2m0 16v2M4.93 4.93l1.42 1.42m11.3 11.3 1.42 1.42M2 12h2m16 0h2M4.93 19.07l1.42-1.42m11.3-11.3 1.42-1.42"/></svg>'
      : '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M20.8 13.2A8.8 8.8 0 0 1 10.8 3.2 8.9 8.9 0 1 0 20.8 13.2Z"/><path d="M16.5 3.5v3m-1.5-1.5h3"/></svg>';
    if (button.classList.contains('theme-toggle-login')) {
      const label = document.createElement('span');
      label.className = 'theme-toggle-label';
      label.textContent = dark ? 'Light mode' : 'Dark mode';
      button.append(label);
    }
  }

  function mountToggle() {
    let button = document.getElementById('themeToggle');
    if (!button) {
      button = document.createElement('button');
      button.id = 'themeToggle';
      button.type = 'button';
      button.className = 'theme-toggle';
    }
    if (document.getElementById('loginForm')) button.classList.add('theme-toggle-login');
    if (!button.dataset.themeBound) {
      button.addEventListener('click', () => saveTheme(root.dataset.theme === 'dark' ? 'light' : 'dark'));
      button.dataset.themeBound = 'true';
    }

    let actions = document.querySelector('.top-actions, .top-links, #accountBar');
    if (!actions) {
      const topInner = document.querySelector('header .top-inner');
      const backLink = topInner?.querySelector(':scope > a.btn');
      if (topInner && backLink) {
        actions = document.createElement('div');
        actions.className = 'top-links theme-actions';
        backLink.replaceWith(actions);
        actions.append(backLink);
        topInner.append(actions);
      }
    }

    if (actions && !button.isConnected) {
      actions.prepend(button);
    } else if (!actions && !button.isConnected) {
      button.classList.add('theme-toggle-floating');
      document.body.append(button);
    }
    updateToggle();
  }

  function closeOtherMenus(openMenu) {
    document.querySelectorAll('details[open]').forEach((menu) => {
      if (menu !== openMenu) menu.open = false;
    });
  }

  root.dataset.theme = readTheme();
  mountToggle();
  document.querySelectorAll('details').forEach((menu) => {
    menu.addEventListener('toggle', () => {
      if (menu.open) closeOtherMenus(menu);
    });
  });
  document.addEventListener('click', (event) => {
    const summary = event.target.closest('details > summary');
    if (summary) {
      closeOtherMenus(summary.parentElement);
      return;
    }
    if (event.target.closest('details')) return;
    closeOtherMenus(null);
  }, true);
})();
