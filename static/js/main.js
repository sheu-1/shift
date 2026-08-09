function positionNowMarker() {
  const marker = document.getElementById("timeline-now");
  if (!marker) return;

  const now = new Date();
  const minutesIntoDay = now.getHours() * 60 + now.getMinutes();
  const percent = (minutesIntoDay / 1440) * 100;

  marker.style.left = percent + "%";
}

positionNowMarker();
setInterval(positionNowMarker, 60000);

// Settings page — capacity stepper
function stepCapacity(inputId, delta) {
  const input = document.getElementById(inputId);
  if (!input) return;
  const current = parseInt(input.value, 10) || 1;
  const next = Math.max(1, current + delta);
  input.value = next;
  updateTotalCapacity();
}

function updateTotalCapacity() {
  const totalEl = document.getElementById("total-capacity");
  if (!totalEl) return;
  const inputs = document.querySelectorAll(".capacity-input");
  let sum = 0;
  inputs.forEach(function(inp) { sum += parseInt(inp.value, 10) || 0; });
  totalEl.textContent = sum;
}

// Wire up live total on settings page
document.addEventListener("DOMContentLoaded", function () {
  document.querySelectorAll(".capacity-input").forEach(function (inp) {
    inp.addEventListener("input", updateTotalCapacity);
  });
});
