document.querySelector('.navtoggle')?.addEventListener('click', () => {
  document.body.classList.toggle('nav-open');
});
// 画像クリックで等倍表示（スクリーンショットは縮小して載せているため）
document.querySelectorAll('figure img').forEach((img) => {
  img.addEventListener('click', () => img.closest('figure').classList.toggle('zoom'));
});
