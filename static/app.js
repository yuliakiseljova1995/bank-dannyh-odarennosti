document.addEventListener("DOMContentLoaded", () => {
  const root = document.documentElement;
  const storedTheme = localStorage.getItem("rd-theme");
  if (storedTheme) root.dataset.theme = storedTheme;
  if (!root.dataset.theme) root.dataset.theme = window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light";
  const themeButton = document.querySelector("[data-theme-toggle]");
  const updateThemeLabel = () => {
    if (!themeButton) return;
    const dark = root.dataset.theme === "dark";
    themeButton.querySelector("span")?.replaceChildren(document.createTextNode(dark ? "☀" : "◐"));
    themeButton.setAttribute("aria-label", dark ? "Включить светлую тему" : "Включить темную тему");
    themeButton.setAttribute("aria-pressed", String(dark));
  };
  themeButton?.addEventListener("click", () => {
    const next = root.dataset.theme === "dark" ? "light" : "dark";
    root.dataset.theme = next;
    localStorage.setItem("rd-theme", next);
    updateThemeLabel();
  });
  updateThemeLabel();

  const menuButton = document.querySelector("[data-menu-toggle]");
  const nav = document.querySelector(".nav");
  const closeMenu = () => { nav?.classList.remove("open"); menuButton?.setAttribute("aria-expanded", "false"); };
  menuButton?.addEventListener("click", () => {
    const open = !nav?.classList.contains("open");
    nav?.classList.toggle("open", open);
    menuButton.setAttribute("aria-expanded", String(open));
  });
  document.addEventListener("click", event => { if (nav?.classList.contains("open") && !nav.contains(event.target) && !menuButton.contains(event.target)) closeMenu(); });
  document.addEventListener("keydown", event => { if (event.key === "Escape") closeMenu(); });

  document.querySelectorAll(".flash").forEach(flash => {
    const close = flash.querySelector(".flash-close");
    close?.addEventListener("click", () => flash.classList.add("is-hidden"));
  });

  const notify = (message, type = "success") => {
    let region = document.querySelector(".toast-region");
    if (!region) {
      region = document.createElement("section");
      region.className = "toast-region floating";
      region.setAttribute("aria-live", "polite");
      document.body.append(region);
    }
    const notice = document.createElement("div");
    notice.className = `flash ${type}`;
    notice.textContent = message;
    region.append(notice);
    window.setTimeout(() => notice.remove(), 3500);
  };

  document.querySelectorAll("table").forEach(table => {
    const headings = Array.from(table.querySelectorAll("thead th"), cell => cell.textContent.trim());
    table.querySelectorAll("tbody tr").forEach(row => {
      row.querySelectorAll("td").forEach((cell, index) => {
        if (!cell.hasAttribute("colspan")) cell.dataset.label = headings[index] || "";
      });
    });
  });

  const chartDataNode = document.querySelector("#dashboard-chart-data");
  if (chartDataNode) {
    const ns = "http://www.w3.org/2000/svg";
    const createSvgNode = (name, attributes = {}) => {
      const node = document.createElementNS(ns, name);
      Object.entries(attributes).forEach(([key, value]) => node.setAttribute(key, value));
      return node;
    };
    const addText = (svg, text, x, y, className, anchor = "start") => {
      const node = createSvgNode("text", { x, y, class: className, "text-anchor": anchor });
      node.textContent = text;
      svg.append(node);
    };
    const renderHorizontalChart = (container, items) => {
      const shown = items.slice(0, 7);
      const height = Math.max(210, shown.length * 48 + 26);
      const svg = createSvgNode("svg", { viewBox: `0 0 680 ${height}`, role: "img" });
      const max = Math.max(1, ...shown.map(item => item.value));
      shown.forEach((item, index) => {
        const y = 18 + index * 48;
        addText(svg, item.label, 0, y + 18, "chart-label");
        const track = createSvgNode("rect", { x: 210, y, width: 400, height: 26, rx: 9, class: "chart-track" });
        const bar = createSvgNode("rect", { x: 210, y, width: Math.max(8, item.value / max * 400), height: 26, rx: 9, class: "chart-bar", tabindex: 0 });
        const title = createSvgNode("title");
        title.textContent = `${item.label}: ${item.value}`;
        bar.append(title);
        svg.append(track, bar);
        addText(svg, String(item.value), 628, y + 18, "chart-value");
      });
      if (!shown.length) addText(svg, "Нет данных по выбранным фильтрам", 340, 108, "chart-empty", "middle");
      container.replaceChildren(svg);
    };
    const renderColumnChart = (container, items) => {
      const svg = createSvgNode("svg", { viewBox: "0 0 680 300", role: "img" });
      const max = Math.max(1, ...items.map(item => item.value));
      const slot = 620 / Math.max(1, items.length);
      items.forEach((item, index) => {
        const width = Math.min(70, slot * .58);
        const barHeight = Math.max(8, item.value / max * 190);
        const x = 30 + index * slot + (slot - width) / 2;
        const bar = createSvgNode("rect", { x, y: 225 - barHeight, width, height: barHeight, rx: 10, class: `chart-bar direction-${item.key}`, tabindex: 0 });
        const title = createSvgNode("title");
        title.textContent = `${item.label}: ${item.value}`;
        bar.append(title);
        svg.append(bar);
        addText(svg, String(item.value), x + width / 2, 219 - barHeight, "chart-value", "middle");
        const label = item.label.length > 18 ? `${item.label.slice(0, 17)}…` : item.label;
        addText(svg, label, x + width / 2, 253, "chart-label small", "middle");
      });
      if (!items.length) addText(svg, "Нет данных", 340, 145, "chart-empty", "middle");
      container.replaceChildren(svg);
    };
    try {
      const data = JSON.parse(chartDataNode.textContent);
      document.querySelectorAll("[data-chart]").forEach(container => {
        if (container.dataset.chart === "schools") renderHorizontalChart(container, data.schools || []);
        if (container.dataset.chart === "directions") renderColumnChart(container, data.directions || []);
        container.previousElementSibling?.remove();
      });
    } catch (_error) {
      document.querySelectorAll(".chart-skeleton").forEach(node => node.remove());
    }
  }
  document.querySelectorAll('form[method="post"], form[method="POST"]').forEach(form => form.addEventListener("submit", event => {
    if (form.dataset.loading === "true") { event.preventDefault(); return; }
    form.dataset.loading = "true";
    const submitter = event.submitter || form.querySelector("button[type=submit],button:not([type])");
    if (submitter) {
      if (submitter.name) {
        const submittedValue = document.createElement("input");
        submittedValue.type = "hidden";
        submittedValue.name = submitter.name;
        submittedValue.value = submitter.value;
        form.append(submittedValue);
      }
      submitter.disabled = true;
      submitter.classList.add("is-loading");
      submitter.dataset.originalText = submitter.textContent;
      submitter.innerHTML = '<span class="spinner" aria-hidden="true"></span> Сохраняем...';
    }
  }));

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
    try {
      await navigator.clipboard.writeText(source.innerText);
      const old = button.textContent; button.textContent = "Скопировано";
      setTimeout(() => button.textContent = old, 1500);
    } catch (_error) {
      notify("Не удалось скопировать текст", "error");
    }
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
      notify(error.message, "error");
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
