/* ============================================
   小书斋编程学习平台 - 认证逻辑 v4
   纯前端静态模式（GitHub Pages 兼容）
   ============================================ */

(function () {
  'use strict';

  var SESSION_KEY = 'xz_session';
  var SECRET_KEY  = '248600';

  function $(id) { return document.getElementById(id); }

  function getSession() {
    try { return JSON.parse(localStorage.getItem(SESSION_KEY)); }
    catch (e) { return null; }
  }

  function setSession(user, isAdmin) {
    localStorage.setItem(SESSION_KEY, JSON.stringify({
      user: user,
      isAdmin: !!isAdmin,
      loggedAt: Date.now()
    }));
  }

  function clearSession() {
    localStorage.removeItem(SESSION_KEY);
  }

  function isLoggedIn() {
    return !!getSession();
  }

  function showError(el, msg) {
    el.textContent = msg;
    el.style.opacity = '1';
  }

  function clearError(el) {
    el.textContent = '';
    el.style.opacity = '0';
  }

  /* ---------- 已登录则跳转 ---------- */
  if (isLoggedIn()) {
    window.location.href = 'dashboard.html';
    return;
  }

  var loginForm    = $('loginForm');
  var registerForm = $('registerForm');

  loginForm.style.display = 'flex';
  registerForm.style.display = 'none';

  /* ---------- 从静态 JSON 加载账号列表 ---------- */
  fetch('data/accounts.json')
    .then(function (r) { return r.json(); })
    .then(function (accounts) {
      if (accounts.length === 0) {
        loginForm.style.display = 'none';
        registerForm.style.display = 'flex';
        var adminChk = $('regIsAdmin');
        if (adminChk) { adminChk.checked = true; }
      } else {
        loginForm.style.display = 'flex';
        registerForm.style.display = 'none';
      }
    })
    .catch(function () {
      loginForm.style.display = 'flex';
      registerForm.style.display = 'none';
    });

  /* ---------- 切换登录 / 注册 ---------- */
  $('toRegister').addEventListener('click', function () {
    loginForm.style.display = 'none';
    registerForm.style.display = 'flex';
    clearError($('loginError'));
  });

  $('toLogin').addEventListener('click', function () {
    registerForm.style.display = 'none';
    loginForm.style.display = 'flex';
    clearError($('regError'));
  });

  /* ---------- 注册逻辑（纯前端 · 临时会话） ---------- */
  registerForm.addEventListener('submit', function (e) {
    e.preventDefault();
    var errEl = $('regError');
    clearError(errEl);

    var key     = $('regKey').value.trim();
    var user    = $('regUser').value.trim();
    var pass    = $('regPass').value;
    var pass2   = $('regPass2').value;

    if (key !== SECRET_KEY) {
      showError(errEl, '注册密钥错误，请联系管理员获取');
      return;
    }
    if (user.length < 2) {
      showError(errEl, '账号至少需要 2 个字符');
      return;
    }
    if (pass.length < 4) {
      showError(errEl, '密码至少需要 4 个字符');
      return;
    }
    if (pass !== pass2) {
      showError(errEl, '两次输入的密码不一致');
      return;
    }

    /* 纯前端模式：检查账号是否已存在 */
    fetch('data/accounts.json')
      .then(function (r) { return r.json(); })
      .then(function (accounts) {
        var exists = accounts.some(function (a) { return a.user === user; });
        if (exists) {
          showError(errEl, '该账号已存在，请直接登录');
          return;
        }

        /* 静态模式：注册为临时会话（仅当前浏览器有效）
           管理员可通过 ppt_manager.py user add 永久添加账号 */
        var adminChk = $('regIsAdmin');
        var isAdmin = adminChk ? adminChk.checked : false;
        setSession(user, isAdmin);
        window.location.href = 'dashboard.html';
      })
      .catch(function () {
        showError(errEl, '网络错误，请稍后重试');
      });
  });

  /* ---------- 登录逻辑（纯前端） ---------- */
  loginForm.addEventListener('submit', function (e) {
    e.preventDefault();
    var errEl = $('loginError');
    clearError(errEl);

    var user = $('loginUser').value.trim();
    var pass = $('loginPass').value;

    fetch('data/accounts.json')
      .then(function (r) { return r.json(); })
      .then(function (accounts) {
        var encodedPass = btoa(pass);
        var found = accounts.find(function (a) {
          return a.user === user && a.pass === encodedPass;
        });
        if (found) {
          setSession(found.user, found.admin);
          window.location.href = 'dashboard.html';
        } else {
          showError(errEl, '账号或密码错误');
        }
      })
      .catch(function () {
        showError(errEl, '网络错误，请稍后重试');
      });
  });

})();
