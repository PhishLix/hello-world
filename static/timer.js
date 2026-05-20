(function () {
  const el = document.querySelector(".timer");
  if (!el) return;
  const clock = document.getElementById("running-clock");
  const start = new Date(el.dataset.start);

  function pad(n) { return n.toString().padStart(2, "0"); }
  function tick() {
    const now = new Date();
    let diff = Math.max(0, Math.floor((now - start) / 1000));
    const h = Math.floor(diff / 3600);
    const m = Math.floor((diff % 3600) / 60);
    const s = diff % 60;
    clock.textContent = `${h}:${pad(m)}:${pad(s)}`;
  }
  tick();
  setInterval(tick, 1000);
})();
