// 出缺席记录查看网站 -- shared front-end behaviour.
// Currently just the "复制全部" (copy all) button on the grouped-by-type
// page; kept as its own file rather than an inline <script> so every page
// can share it via <script src="app.js"> if more interactivity is added.

document.addEventListener('DOMContentLoaded', function () {
  var copyBtn = document.getElementById('copy-btn');
  if (!copyBtn) return;

  copyBtn.addEventListener('click', async function () {
    var ta = document.getElementById('grouped-text');
    var status = document.getElementById('copy-status');
    if (!ta) return;

    try {
      await navigator.clipboard.writeText(ta.value);
    } catch (e) {
      ta.select();
      document.execCommand('copy');
    }

    if (status) {
      status.textContent = '已复制到剪贴板 ✓';
      setTimeout(function () { status.textContent = ''; }, 2000);
    }
  });
});
