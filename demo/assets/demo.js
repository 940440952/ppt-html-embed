let n = 0; const c = document.getElementById('c').getContext('2d');
document.getElementById('b').onclick = () => {
  document.getElementById('n').textContent = ++n;
  c.fillStyle = `hsl(${n * 37 % 360} 70% 55%)`;
  c.fillRect((n * 23) % 280, (n * 11) % 60, 20, 20);
};
