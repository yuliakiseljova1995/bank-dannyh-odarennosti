document.addEventListener("DOMContentLoaded", () => {
  const winner = document.querySelector("#mentor-winner");
  const result = document.querySelector("#mentor-result");
  const toggleWinner = () => result?.classList.toggle("visible", !!winner?.checked);
  winner?.addEventListener("change", toggleWinner);
  toggleWinner();

  const contest = document.querySelector("#contest-name");
  const direction = document.querySelector("#direction");
  let timer;
  contest?.addEventListener("input", () => {
    if (direction.value || contest.value.trim().length < 4) return;
    clearTimeout(timer);
    timer = setTimeout(async () => {
      const response = await fetch(`/api/suggest-direction?contest=${encodeURIComponent(contest.value)}`);
      if (response.ok && !direction.value) direction.value = (await response.json()).direction;
    }, 350);
  });

  document.querySelectorAll("[data-copy]").forEach(button => button.addEventListener("click", async () => {
    const source = document.querySelector(button.dataset.copy);
    await navigator.clipboard.writeText(source.innerText);
    const old = button.textContent; button.textContent = "Скопировано";
    setTimeout(() => button.textContent = old, 1500);
  }));

  document.querySelectorAll("[data-generate-social]").forEach(button => button.addEventListener("click", async () => {
    const old = button.textContent;
    button.disabled = true;
    button.textContent = "Генерация...";
    try {
      const response = await fetch(button.dataset.url, {
        method: "POST",
        headers: { "X-CSRF-Token": button.dataset.csrf },
      });
      const data = await response.json();
      if (!response.ok) throw new Error(data.error || "Не удалось создать текст");
      document.querySelector("#social-text").innerText = data.text;
      button.textContent = "Текст готов";
      setTimeout(() => button.textContent = old, 1800);
    } catch (error) {
      window.alert(error.message);
      button.textContent = old;
    } finally {
      button.disabled = false;
    }
  }));

  document.querySelectorAll("[data-response-form]").forEach(form => form.addEventListener("submit", event => {
    form.querySelectorAll("button").forEach(button => {
      button.disabled = true;
      if (button === event.submitter) button.textContent = "Сохраняем...";
    });
  }));
});
