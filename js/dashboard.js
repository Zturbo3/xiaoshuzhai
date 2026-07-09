/* ============================================
   小书斋编程学习平台 - 仪表盘交互 v2
   ============================================ */

(function () {
  'use strict';

  var SESSION_KEY = 'xz_session';

  /* ---------- 登录状态检查 ---------- */
  function getSession() {
    try { return JSON.parse(localStorage.getItem(SESSION_KEY)); }
    catch (e) { return null; }
  }

  var session = getSession();
  if (!session) {
    window.location.href = 'index.html';
    return;
  }

  /* ---------- 显示用户信息 ---------- */
  var nameEl = document.getElementById('currentUser');
  if (nameEl) {
    nameEl.textContent = session.user;
  }

  /* 管理员标识 */
  var badgeEl = document.getElementById('adminBadge');
  if (badgeEl && session.isAdmin) {
    badgeEl.style.display = 'inline';
  }

  /* ---------- 退出登录 ---------- */
  var btnLogout = document.getElementById('btnLogout');
  if (btnLogout) {
    btnLogout.addEventListener('click', function () {
      localStorage.removeItem(SESSION_KEY);
      window.location.href = 'index.html';
    });
  }

  /* ---------- 卡片点击导航 ---------- */
  var routes = {
    scratch: 'scratch.html',
    python:  'python.html',
    cpp:     'cpp.html'
  };

  var cards = document.querySelectorAll('.card');
  cards.forEach(function (card) {
    card.addEventListener('click', function () {
      var type = card.getAttribute('data-card');
      var target = routes[type];
      if (target) {
        card.style.transform = 'translateY(-12px) scale(0.97)';
        setTimeout(function () {
          card.style.transform = '';
          window.location.href = target;
        }, 200);
      }
    });
  });

})();
